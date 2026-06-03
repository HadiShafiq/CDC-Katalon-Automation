import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webui.common.WebUiCommonHelper
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.Keys
import org.openqa.selenium.WebDriver
import org.openqa.selenium.WebElement
import org.openqa.selenium.chrome.ChromeDriver
import org.openqa.selenium.chrome.ChromeOptions

import java.io.FileInputStream
import java.io.FileOutputStream
import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat
import java.util.Arrays

import org.apache.poi.xssf.usermodel.XSSFWorkbook

/* =========================================================
 * HELPER SECTION
 * Purpose:
 * - common reusable functions
 * - grouped by type for easier maintenance
 * - logic unchanged, only rearranged and commented
 * ========================================================= */


/* =========================================================
 * 1) BASIC DATA / CONVERSION HELPER
 * ========================================================= */

/**
 * Convert Excel / CSV value to int safely.
 * Example supported values:
 * "0", 0, 0.0, "1.0"
 */
int toInt(def v, int defaultVal = 0) {
	if (v == null) return defaultVal
	return new BigDecimal(v.toString().trim()).intValue()
}


/* =========================================================
 * 2) PAGE / OVERLAY WAIT HELPER
 * ========================================================= */

/**
 * Wait until PrimeFaces / block UI overlay disappears.
 * Use this after click, dropdown selection, save, popup action, etc.
 */
def waitBlockUI(int timeout = 30) {
	TestObject blockUI = new TestObject('blockUI')
	blockUI.addProperty("xpath", ConditionType.EQUALS,
		"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
	)

	if (WebUI.verifyElementPresent(blockUI, 1, FailureHandling.OPTIONAL)) {
		WebUI.waitForElementNotVisible(blockUI, timeout, FailureHandling.OPTIONAL)
	}
}


/* =========================================================
 * 3) LIGHTWEIGHT ELEMENT WAIT / ACTION HELPERS
 * ========================================================= */

/**
 * Wait until element is visible.
 */
def wVisible(TestObject obj, int timeout = 1) {
	waitBlockUI(Math.min(timeout, 1))
	WebUI.waitForElementVisible(obj, timeout, FailureHandling.STOP_ON_FAILURE)
}

/**
 * Wait until element is clickable.
 * This first ensures element is visible.
 */
def wClickable(TestObject obj, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.waitForElementClickable(obj, timeout, FailureHandling.STOP_ON_FAILURE)
}

/**
 * Click element with:
 * - wait
 * - scroll
 * - small retry if first click fails
 */
def c(TestObject obj, int timeout = 1) {
	for (int i=0; i<2; i++) {
		try {
			wClickable(obj, timeout)
			WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
			WebUI.click(obj)
			waitBlockUI(1)
			return
		} catch (Exception e) {
			WebUI.delay(0.3)
		}
	}
	// last try
	wClickable(obj, timeout)
	WebUI.click(obj)
	waitBlockUI(1)
}

/**
 * Double click element with wait.
 */
def dc(TestObject obj, int timeout = 1) {
	try {
		wClickable(obj, timeout)
		WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
		WebUI.doubleClick(obj, FailureHandling.OPTIONAL)
		waitBlockUI(1)
	} catch (Exception e) {
		WebUI.doubleClick(obj, FailureHandling.OPTIONAL)
		waitBlockUI(1)
	}
}

/**
 * Set text with wait.
 * Best for normal text field.
 * For sensitive formatted field, use JS/manual typing separately.
 */
def t(TestObject obj, def value, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
	WebUI.setText(obj, (value == null ? "" : value.toString()))
}

/**
 * Upload file with wait.
 */
def up(TestObject obj, String filePath, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.uploadFile(obj, filePath)
	waitBlockUI(1)
}


/* =========================================================
 * 4) PRIMEFACES DROPDOWN HELPERS
 * ========================================================= */

/**
 * Open PrimeFaces dropdown panel.
 * Used for dropdown that is not a real <select>.
 */
def openPFDropdown(TestObject triggerObj) {

	TestObject panelOpen = new TestObject('pfPanelOpen')
	panelOpen.addProperty("xpath", ConditionType.EQUALS,
		"//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]"
	)

	c(triggerObj, 20)
	WebUI.delay(0.3)

	if (!WebUI.waitForElementVisible(panelOpen, 1, FailureHandling.OPTIONAL)) {
		c(triggerObj, 20)
		WebUI.delay(0.3)
		WebUI.waitForElementVisible(panelOpen, 3, FailureHandling.OPTIONAL)
	}
}

/**
 * Click PrimeFaces dropdown option by 0-based index.
 * Example:
 * 0 = first option
 * 1 = second option
 */
def clickPFOptionByIndex(int index0) {
	TestObject opt = new TestObject("pfOpt_" + index0)
	opt.addProperty("xpath", ConditionType.EQUALS,
		"(//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[contains(@class,'ui-selectonemenu-item')])[${index0 + 1}]"
	)

	c(opt, 20)
	WebUI.delay(0.2)
	waitBlockUI(20)
}

/**
 * Universal dropdown select by index.
 * Supports:
 * - real <select>
 * - PrimeFaces custom dropdown
 */
def selectDropdownByIndex(TestObject dropdownObj, def indexFromData) {

	int idx0 = toInt(indexFromData) // data already 0-based

	for (int attempt = 0; attempt < 3; attempt++) {
		try {
			wVisible(dropdownObj, 20)
			def el = WebUiCommonHelper.findWebElement(dropdownObj, 10)
			String tag = el.getTagName()

			if (tag != null && tag.equalsIgnoreCase("select")) {
				WebUI.selectOptionByIndex(dropdownObj, idx0)
				WebUI.delay(0.3)
				waitBlockUI(20)
			} else {
				openPFDropdown(dropdownObj)
				clickPFOptionByIndex(idx0)
			}
			return
		} catch (org.openqa.selenium.StaleElementReferenceException e) {
			WebUI.delay(0.5)
		}
	}

	assert false : "❌ Dropdown failed (stale/DOM refresh): " + dropdownObj.getObjectId()
}


/* =========================================================
 * 5) TABLE / POPUP INPUT HELPERS
 * ========================================================= */

/**
 * Set zone quantity by popup row index.
 * Used for popup table:
 * specZoneQtyTbl:{rowIndex}:zoneQty
 */
def setZoneQtyByRow = { int rowIndex, String qtyValue ->
	String xpath = "//div[contains(@class,'ui-dialog')]//input[contains(@id,'specZoneQtyTbl:${rowIndex}:zoneQty')]"

	TestObject qtyObj = new TestObject("zoneQty_" + rowIndex)
	qtyObj.addProperty("xpath", ConditionType.EQUALS, xpath)

	WebUI.waitForElementVisible(qtyObj, 20)
	WebElement qtyEl = WebUiCommonHelper.findWebElement(qtyObj, 20)

	WebUI.executeJavaScript(
		"""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """,
		Arrays.asList(qtyEl, qtyValue)
	)

	waitBlockUI(30)
	WebUI.delay(0.5)
}

/**
 * Set order quantity/ rate by popup row index.
 * Used for popup table:
 * specAnswerTbl:{rowIndex}:ratePerUomAns
 */
def setOrderedQtyByRow = { int rowIndex, String qtyValue ->

    String xpath = "//div[contains(@class,'ui-dialog')]//tr[@data-ri='" + rowIndex + "']//input[contains(@name,'orderedQty')]"

    TestObject qtyObj = new TestObject("orderedQty_" + rowIndex)
    qtyObj.addProperty("xpath", ConditionType.EQUALS, xpath)

    WebUI.waitForElementVisible(qtyObj, 20)
    WebElement qtyEl = WebUiCommonHelper.findWebElement(qtyObj, 20)

    WebUI.executeJavaScript(
        """
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """,
        Arrays.asList(qtyEl, qtyValue)
    )

    waitBlockUI(30)
    WebUI.delay(0.5)
}


/* =========================================================
 * 6) SPECIAL RADIO / OPTION HELPER
 * ========================================================= */

/**
 * Click procurement type radio button by business option number.
 * Mapping:
 * 1 = first radio
 * 2 = second radio
 * etc.
 */
def clickProcurementType(int option) {
	String xpath = "//*[@id='_Catalogue_WAR_NGePportlet_:form:procurementType:${option - 1}']"

	TestObject obj = new TestObject("procurementType_" + option)
	obj.addProperty("xpath", ConditionType.EQUALS, xpath)

	WebUI.waitForElementVisible(obj, 20)
	WebUI.waitForElementClickable(obj, 20)

	WebUI.executeJavaScript(
		"arguments[0].click();",
		Arrays.asList(WebUiCommonHelper.findWebElement(obj, 20))
	)

	waitBlockUI(20)
}


/* =========================================================
 * 7) CALENDAR PICKER DATE
 * ========================================================= */

def pickDate(String yyyyMmDd) {

	TestObject dp = new TestObject('dp')
	dp.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='ui-datepicker-div' and not(contains(@style,'display: none'))]"
	)
	WebUI.waitForElementVisible(dp, 20)

	def parts = yyyyMmDd.split('-')
	int targetYear = parts[0] as int
	int targetMonthIndex = (parts[1] as int) - 1
	String targetDay = String.valueOf(parts[2] as int)

	TestObject nextBtn = new TestObject('dpNext')
	nextBtn.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='ui-datepicker-div']//a[contains(@class,'ui-datepicker-next')]"
	)

	TestObject prevBtn = new TestObject('dpPrev')
	prevBtn.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='ui-datepicker-div']//a[contains(@class,'ui-datepicker-prev')]"
	)

	int guard = 0
	while (guard < 48) {

		TestObject targetDayObj = new TestObject("targetDay_${targetYear}_${targetMonthIndex}_${targetDay}")
		targetDayObj.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='ui-datepicker-div']//td[@data-year='${targetYear}' and @data-month='${targetMonthIndex}' " +
			"and not(contains(@class,'ui-state-disabled'))]//a[normalize-space()='${targetDay}']"
		)

		if (WebUI.verifyElementPresent(targetDayObj, 1, FailureHandling.OPTIONAL)) {
			WebUI.waitForElementClickable(targetDayObj, 20)
			WebUI.click(targetDayObj)
			return
		}

		// kalau target belum ada, click next dulu
		WebUI.click(nextBtn)
		WebUI.delay(0.3)
		guard++
	}

	WebUI.takeScreenshot()
	assert false : "Date not found in datepicker: " + yyyyMmDd
}

/* =========================================================
 * 8) BROWSER SETUP
 * ========================================================= */
// USE ENVIRONMENT VARIABLE 	
String chromeBinary = System.getenv("CHROME_BINARY_PATH")
String chromeDriverPath = System.getenv("CHROME_DRIVER_PATH")

System.setProperty("webdriver.chrome.driver", chromeDriverPath)

String userDataDir = Files.createTempDirectory("katalon-cft").toString()

ChromeOptions options = new ChromeOptions()
options.setBinary(chromeBinary)

// Bypass browser security popup
options.setAcceptInsecureCerts(true)

// Disable HTTPS-first strict behavior
options.addArguments("--disable-features=HttpsFirstBalancedModeAutoEnable,HttpsUpgrades")

// Browser profile options
options.addArguments("--guest")
// options.addArguments("--incognito")
options.addArguments("--user-data-dir=" + userDataDir)
options.addArguments("--disable-features=PasswordLeakDetection,PasswordManagerOnboarding")
options.addArguments("--disable-save-password-bubble")
options.addArguments("--disable-notifications")
options.addArguments("--no-first-run")
options.addArguments("--no-default-browser-check")
options.addArguments("--remote-allow-origins=*")

// Browser preferences
Map<String, Object> prefs = new HashMap<>()
prefs.put("credentials_enable_service", false)
prefs.put("profile.password_manager_enabled", false)
prefs.put("profile.default_content_setting_values.notifications", 2)
options.setExperimentalOption("prefs", prefs)

// Launch browser
WebDriver driver = new ChromeDriver(options)
DriverFactory.changeWebDriver(driver)

/* =========================
 * OPEN APPLICATION
 * Purpose:
 * - open NGeP SIT portal
 * - maximize browser
 * - wait initial page load
 * ========================= */
WebUI.navigateToUrl('http://ngepsit.eperolehan.com.my/home')
WebUI.maximizeWindow()
waitBlockUI(20)

/* =========================
 * LANGUAGE
 * Purpose:
 * - switch system language to English
 * ========================= */
wVisible(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 20)
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)
waitBlockUI(20)
WebUI.delay(0.5)
WebUI.delay(1)

/* =========================
 * LOGIN
 * Purpose:
 * - open login form
 * - enter username and password
 * - submit login
 * ========================= */
c(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Right Top Menu Login'), 20)
WebUI.delay(0.5)

t(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Username'), Username, 20)
WebUI.delay(0.5)

t(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Password'), Password, 20)
WebUI.delay(0.5)

c(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Submit Username and Password'), 20)
waitBlockUI(30)
WebUI.delay(0.5)

/* =========================
 * LANGUAGE
 * Purpose:
 * -Change language inside dashboard
 * ========================= */
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

//TaskList
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Click Task List'))

c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/MyTask_Tasklist_Dropdown'))

//Input Document Number
t(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Input Document Number'), 
    Document_Number)

c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Search TaskList'))

//Click TaskList Description
c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Click TaskList Description'))

// =========================
// General - Tick 
// =========================
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Checkbox'))
waitBlockUI(20)
WebUI.delay(0.5)

// =========================
// Delivery Address & Item - PTJ
// =========================
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Menu Delivery'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Button PTJ Address'))
selectDropdownByIndex(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Dropdown Searching'), Address)
t(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Address Input'), Address_Input)
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Search Button'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Checkbox Address'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Select Button'))
waitBlockUI(20)
WebUI.delay(0.5)

// =========================
// Delivery Address & Item - Ordered Quantity
// =========================
int loopCount = 1

for (int i = 0; i < loopCount; i++) {

    String xpath = "(//input[contains(@name,'orderedQty')])[" + (i + 1) + "]"

    TestObject orderedQuantityField = new TestObject("orderedQuantityField_" + i)
    orderedQuantityField.addProperty("xpath", ConditionType.EQUALS, xpath)

    WebUI.comment("Fill Ordered Quantity row #" + (i + 1))

    WebUI.waitForElementVisible(orderedQuantityField, 20)

    WebElement el = WebUiCommonHelper.findWebElement(orderedQuantityField, 20)

    // 🔥 1. CLEAR VALUE FIRST
    WebUI.executeJavaScript(
        """
        arguments[0].value = '';
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """,
        Arrays.asList(el)
    )

    WebUI.delay(0.5)

    // 🔥 2. SET NEW VALUE
    t(orderedQuantityField, Ordered_Quantity, 20)

    waitBlockUI(20)
    WebUI.delay(1)

    // 🔥 3. TRIGGER TAB (force JSF update)
    WebUI.sendKeys(orderedQuantityField, Keys.chord(Keys.TAB))

    waitBlockUI(20)
    WebUI.delay(1)
}

// =========================
// Charge Line Assigment - TICK
// =========================
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Menu Charge Line Assignment'))
int loopCountA = 1

for (int i = 0; i < loopCountA; i++) {

    WebUI.comment("Tick row #" + (i + 1))

    String chkXpath = "(//input[contains(@id,'chargeLineTbl') and contains(@type,'checkbox')])[" + (i + 1) + "]"

    TestObject chkObj = new TestObject("chk_" + i)
    chkObj.addProperty("xpath", ConditionType.EQUALS, chkXpath)

    WebUI.waitForElementClickable(chkObj, 2)

    WebElement chkEl = WebUiCommonHelper.findWebElement(chkObj, 2)

    if (!chkEl.isSelected()) {
        WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(chkEl))
    }

    WebUI.delay(0.5)
}

selectDropdownByIndex(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Transaction Indicator'), Transaction_Indicator)

//VOT
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Click VOT 1'))
selectDropdownByIndex(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Dropdown Searching'), Listing_VOT)
t(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Listing-VOT_Input'), VOT_Search)
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Search Button'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Choose Result'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Select Button Search'))

//Program/Activity
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Click Program_Activity'))
selectDropdownByIndex(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Dropdown Searching'), Listing_Program)
t(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Listing-VOT_Input'), Program_Search)
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Choose Result'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Select Button Search'))

//Account Code
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Account Code'))

// =========================
// Input Description Search
// =========================
TestObject descriptionInput = findTestObject(
	'Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Description_Input'
)

String descriptionValue = Description_Search.toString().trim()

WebUI.waitForElementVisible(descriptionInput, 20)
WebUI.waitForElementClickable(descriptionInput, 20)

WebUI.click(descriptionInput)
WebUI.delay(0.5)

WebUI.sendKeys(descriptionInput, Keys.chord(Keys.CONTROL, 'a'))
WebUI.delay(0.3)

WebUI.sendKeys(descriptionInput, Keys.chord(Keys.BACK_SPACE))
WebUI.delay(0.5)

for (char ch : descriptionValue.toCharArray()) {
	WebUI.sendKeys(descriptionInput, ch.toString())
	WebUI.delay(0.15)
}

waitBlockUI(10)
WebUI.delay(1)

String finalDescription = WebUI.getAttribute(descriptionInput, 'value')
println("Final Description Search = " + finalDescription)

c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Button Search Code'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Choose Result 1'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Select Button Search 1'))
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Click Button Update'))

// ==================================
// Approver List - Choose Approver
// ==================================
c(findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Menu Approver List'))
TestObject approverGroup = findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Approver Group Dropdown')
TestObject approverName  = findTestObject('Object Repository/DLOA/9. DLOA Supplier/Purchase Request/Approver Name Dropdown')
			
// Select Approver Group
selectDropdownByIndex(approverGroup, 11)
waitBlockUI(20)
WebUI.delay(2)
			
// Wait until Approver Name dropdown is ready
wVisible(approverName, 1)
WebUI.waitForElementClickable(approverName, 1)
WebUI.scrollToElement(approverName, 1)
			
// Select Approver Name
selectDropdownByIndex(approverName, 5)
WebUI.delay(1)


// ========================
// SUBMIT BUTTON
//=========================
c(findTestObject('Object Repository/FD and Agreement/FD Application/Approver Setting/Submit Button'))
waitBlockUI(10)
WebUI.delay(0.5)

// ===== 1) Wait loader/blockUI gone (PrimeFaces common) =====
TestObject blockUI = new TestObject('blockUI')
blockUI.addProperty("xpath", ConditionType.EQUALS,
	"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
)

if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
	WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
}

// ===== 2) Wait success message (global text; RN number changes) =====
TestObject msgObj = new TestObject('msg_PR_saved')
msgObj.addProperty("xpath", ConditionType.EQUALS,
	"//span[contains(@class,'ui-messages-info-detail') and " +
	"contains(.,'Purchase Request') and contains(.,'is successfully submitted.')]"  
)

WebUI.waitForElementVisible(msgObj, 30)

// Wait until message text contains "PR"
String msg = ""
for (int i = 0; i < 2; i++) {
	msg = WebUI.getText(msgObj, FailureHandling.OPTIONAL)
	if (msg != null && msg.contains("PR")) break
	WebUI.delay(1)
}

msg = (msg == null) ? "" : msg.trim()
WebUI.comment("Message: " + msg)

// ===== 3) Extract RN number dynamically =====
def matcher = (msg =~ /(PR\d+)/)   // e.g. RN260000000001152
String prNo = matcher.find() ? matcher.group(1) : ""

if (prNo == "") {
	WebUI.takeScreenshot()
	assert false : "❌ PR number not found. Message was: " + msg
}
WebUI.comment("✅ Captured PR No: " + prNo)

// ===== 4) Append to SAME Excel file (no timestamp file) =====
String baseDir = System.getProperty("user.home") + "/Desktop/PrepDataFileNumber"
new File(baseDir).mkdirs() //AUTO-CREATE FOLDER
String filePath = baseDir + "/DLOA_PURCHASE_REQUEST_2026.xlsx"
String now = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date())

def path = Paths.get(filePath)
XSSFWorkbook wb
def sheet
FileInputStream fis = null

if (Files.exists(path)) {
	fis = new FileInputStream(filePath)
	wb = new XSSFWorkbook(fis)
	sheet = wb.getSheet("Result")
	if (sheet == null) sheet = wb.createSheet("Result")
} else {
	wb = new XSSFWorkbook()
	sheet = wb.createSheet("Result")

	def header = sheet.createRow(0)
	header.createCell(0).setCellValue("DateTime")
	header.createCell(1).setCellValue("PR No")
	header.createCell(2).setCellValue("Message")
}

// Close input stream to avoid Excel file lock
if (fis != null) fis.close()

// Next empty row
int nextRow = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : sheet.getLastRowNum() + 1
def row = sheet.createRow(nextRow)

row.createCell(0).setCellValue(now)
row.createCell(1).setCellValue(prNo)
row.createCell(2).setCellValue(msg)

// Save back to SAME file
FileOutputStream fos = new FileOutputStream(filePath)
wb.write(fos)
fos.close()
wb.close()

WebUI.comment("✅ Appended to Excel: " + filePath)

/* =========================
 * SIGN OUT
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))
WebUI.waitForPageLoad(20)
WebUI.closeBrowser()