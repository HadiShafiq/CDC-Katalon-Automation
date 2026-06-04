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
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.By
import org.openqa.selenium.WebElement
import com.kms.katalon.core.webui.driver.DriverFactory

/* =========================
 * HELPERS TEST
 * Purpose:
 * - reusable utility for conversion, wait, click, text input, upload
 * - keep script stable without changing main flow
 * ========================= */

// Convert excel/csv value to int safely: "0", 0, 0.0, "1.0"
int toInt(def v, int defaultVal = 0) {
	if (v == null) return defaultVal
	return new BigDecimal(v.toString().trim()).intValue()
}

// PrimeFaces overlay wait
def waitBlockUI(int timeout = 30) {
	TestObject blockUI = new TestObject('blockUI')
	blockUI.addProperty("xpath", ConditionType.EQUALS,
		"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
	)

	if (WebUI.verifyElementPresent(blockUI, 1, FailureHandling.OPTIONAL)) {
		WebUI.waitForElementNotVisible(blockUI, timeout, FailureHandling.OPTIONAL)
	}
}

/* ---------- Lightweight wait wrappers ---------- */
// wait until element visible
def wVisible(TestObject obj, int timeout = 1) {
	waitBlockUI(Math.min(timeout, 1))
	WebUI.waitForElementVisible(obj, timeout, FailureHandling.STOP_ON_FAILURE)
}

// wait until element clickable
def wClickable(TestObject obj, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.waitForElementClickable(obj, timeout, FailureHandling.STOP_ON_FAILURE)
}

// click with wait + tiny retry
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

// double click with wait
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

// setText with wait
def t(TestObject obj, def value, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
	WebUI.setText(obj, (value == null ? "" : value.toString()))
}

/* =========================
 * HELPERS for zone quantity
 * ========================= */
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

/* =========================
 * HELPERS for zone quantity
 * ========================= */

def setUnitPriceByRow = { int rowIndex, String unitPriceValue ->
	String xpath = "//div[contains(@class,'ui-dialog')]//input[contains(@id,'specAnswerTbl:${rowIndex}:ratePerUomAns')]"

	TestObject priceObj = new TestObject("unitPrice_" + rowIndex)
	priceObj.addProperty("xpath", ConditionType.EQUALS, xpath)

	WebUI.waitForElementVisible(priceObj, 20)
	WebElement priceEl = WebUiCommonHelper.findWebElement(priceObj, 20)

	WebUI.executeJavaScript(
		"""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """,
		Arrays.asList(priceEl, unitPriceValue)
	)

	waitBlockUI(30)
	WebUI.delay(0.5)
}
// upload with wait
def up(TestObject obj, String filePath, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.uploadFile(obj, filePath)
	waitBlockUI(1)
}

/* =========================
 * PRIMEFACES DROPDOWN HELPERS
 * Purpose:
 * - open PrimeFaces dropdown
 * - click option by index
 * - support both PrimeFaces and real select
 * ========================= */
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

// click PrimeFaces option by index (0-based)
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
 * Universal dropdown select by index (SAFE)
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

/* =========================
 * Function: Claim document by Document No
 * ========================= */
def claimDocument(String targetDocNo) {

	boolean found = false

	for (int pageIndex = 1; pageIndex <= 10; pageIndex++) {

		WebUI.comment("Checking page " + pageIndex + " for Document No: " + targetDocNo)

		TestObject targetRow = new TestObject("targetRow_" + targetDocNo)
		targetRow.addProperty("xpath", ConditionType.EQUALS,
			"//tbody[contains(@id,'taskListGroupId_data')]//tr[td[normalize-space()='" + targetDocNo + "']]"
		)

		if (WebUI.verifyElementPresent(targetRow, 3, FailureHandling.OPTIONAL)) {

			TestObject claimBtn = new TestObject("claimBtn_" + targetDocNo)
			claimBtn.addProperty("xpath", ConditionType.EQUALS,
				"//tbody[contains(@id,'taskListGroupId_data')]//tr[td[normalize-space()='" + targetDocNo + "']]//span[normalize-space()='Claim']/ancestor::button"
			)

			WebUI.waitForElementClickable(claimBtn, 20)
			c(claimBtn, 20)
			waitBlockUI(30)
			WebUI.delay(1)

			found = true
			break
		}

		if (pageIndex < 10) {
			int nextIndex = pageIndex + 1

			TestObject nextPage = new TestObject("page_" + nextIndex)
			nextPage.addProperty("xpath", ConditionType.EQUALS,
				"(//span[contains(@class,'ui-paginator-pages')]/span[normalize-space()='" + nextIndex + "'])[1]"
			)

			if (WebUI.verifyElementPresent(nextPage, 5, FailureHandling.OPTIONAL)) {
				WebUI.scrollToElement(nextPage, 3)
				WebUI.click(nextPage)
				waitBlockUI(30)
				WebUI.delay(1)
			} else {
				break
			}
		}
	}

	return found
}

/* ===========================================================================
 * Function: Untuk Side Menu Schedule | Performance Bond | Payment Tracking
 * =========================================================================== */
def clickSideMenuIfExists(String objectPath) {
	
		TestObject menuObj = findTestObject(objectPath)
	
		if (WebUI.waitForElementClickable(menuObj, 5, FailureHandling.OPTIONAL)) {
			c(menuObj)
			waitBlockUI(20)
			return true
		}
	
		WebUI.comment("Skip: menu not available -> " + objectPath)
		return false
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

//Bypass security pop up for google chrome 
options.setAcceptInsecureCerts(true)

options.addArguments("--disable-features=HttpsFirstBalancedModeAutoEnable,HttpsUpgrades")

options.addArguments("--guest")
//options.addArguments("--incognito")
options.addArguments("--user-data-dir=" + userDataDir)
options.addArguments("--disable-features=PasswordLeakDetection,PasswordManagerOnboarding")
options.addArguments("--disable-save-password-bubble")
options.addArguments("--disable-notifications")
options.addArguments("--no-first-run")
options.addArguments("--no-default-browser-check")
options.addArguments("--remote-allow-origins=*")

Map<String, Object> prefs = new HashMap<>()
prefs.put("credentials_enable_service", false)
prefs.put("profile.password_manager_enabled", false)
prefs.put("profile.default_content_setting_values.notifications", 2)
options.setExperimentalOption("prefs", prefs)


WebDriver driver = new ChromeDriver(options)
DriverFactory.changeWebDriver(driver)

/* ============================
 * String for Upload Files
 * ============================ */
String uploadFilePath = System.getProperty("user.dir") + "/TestData/UploadFiles/File_pdf_for_testing.pdf"

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
 * - enter username and password 123
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
 * CHANGE LANGUAGE
 * Purpose:
 * - Change Language at Dashboard
 * ========================= */
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

/* =========================
 * Tasklist MyTask
 * Purpose:
 * - Serach Application No
 * =========================*/ 
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Click Task List'))

c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/MyTask_Tasklist_Dropdown'))

//Input Document Number
t(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Input Document Number'),Document_Number)

c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Search TaskList'))
waitBlockUI(20)

TestObject firstRow = findTestObject('Object Repository/DP - Add To Cart/Fulfilment Received Note/Click Tasklist')
WebUI.waitForElementVisible(firstRow, 20)
WebUI.scrollToElement(firstRow, 10)

c(findTestObject('Object Repository/DP - Add To Cart/Fulfilment Received Note/Click Tasklist'))

//Acknowledge Officer
selectDropdownByIndex(findTestObject('Object Repository/DP - Add To Cart/Fulfilment Received Note/Click Dropdown Acknowledge'),ddAcknowledge)
waitBlockUI(20)
WebUI.delay(0.5)

//Menu Item 
c(findTestObject('Object Repository/DP - Add To Cart/Fulfilment Received Note/Click Menu Item'))

//Menu Item
c(findTestObject('Object Repository/DP - Add To Cart/Fulfilment Received Note/Click Menu Supplier Evaluation'))

//Dropdown score
String scoreValue = ddScore

for (int i = 0; i < 5; i++) {

	TestObject dropdown = new TestObject("dropdown_" + i)
	dropdown.addProperty(
		"xpath",
		ConditionType.EQUALS,
		"//select[contains(@name,'suppEvalutionDt:" + i + ":score')]"
	)

	WebUI.waitForElementVisible(dropdown, 10)
	WebUI.waitForElementClickable(dropdown, 10)

	WebUI.selectOptionByValue(dropdown, scoreValue, false)

	WebUI.delay(0.5) // 🔥 penting untuk slow UI sync
	
}

/* ========================
 * SUBMIT BUTTON
 * ========================*/
c(findTestObject('Object Repository/FD and Agreement/FD Application/Approver Setting/Submit Button'))
waitBlockUI(10)
WebUI.delay(0.5)

/* ======================================
 * SUCCESS MESSAGE - After click submit
 * ====================================== */

TestObject blockUI = new TestObject('blockUI')
blockUI.addProperty("xpath", ConditionType.EQUALS,
	"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
)

if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
	WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
}

// ambil ANY message
TestObject msgObj = new TestObject('msg_any')
msgObj.addProperty("xpath", ConditionType.EQUALS,
	"//*[contains(@class,'ui-messages-info-detail') or contains(@class,'ui-messages-warn-detail') or contains(@class,'ui-messages-error-detail')]"
)

WebUI.waitForElementVisible(msgObj, 30)

String msg = WebUI.getText(msgObj, FailureHandling.STOP_ON_FAILURE)
msg = (msg == null) ? "" : msg.trim()

WebUI.comment("Message: " + msg)

// extract FN number
def matcher = (msg =~ /(FN\d+)/)
String fnNo = matcher.find() ? matcher.group(1) : ""

if (fnNo == "") {
	WebUI.takeScreenshot()
	assert false : "❌ FN number not found. Message was: " + msg
}

WebUI.comment("✅ Captured FN No: " + poNum)

/* =========================
 * EXCEL APPEND 
 * ========================= */

String baseDir  = System.getProperty('user.home') + '/Desktop/PrepDataFileNumber'
String filePath = baseDir + '/Submit_Fulfillment_Received_Note.xlsx'
String now      = new SimpleDateFormat('yyyy-MM-dd HH:mm:ss').format(new Date())

new File(baseDir).mkdirs()

XSSFWorkbook wb
def sheet
FileInputStream fis = null
def path = Paths.get(filePath)

try {
	if (Files.exists(path)) {
		fis = new FileInputStream(filePath)
		wb = new XSSFWorkbook(fis)
		sheet = wb.getSheet('Result') ?: wb.createSheet('Result')
	} else {
		wb = new XSSFWorkbook()
		sheet = wb.createSheet('Result')

		// header
		def header = sheet.createRow(0)
		header.createCell(0).setCellValue('DateTime')
		header.createCell(1).setCellValue('FN No')
		header.createCell(2).setCellValue('Message')
	}

	int nextRow = sheet.getLastRowNum() + 1
	def row = sheet.createRow(nextRow)

	row.createCell(0).setCellValue(now)
	row.createCell(1).setCellValue(fnNo)
	row.createCell(2).setCellValue(msg)

	FileOutputStream fos = new FileOutputStream(filePath)
	wb.write(fos)
	fos.flush()
	fos.close()

} catch (Exception e) {
	WebUI.comment("❌ Gagal menulis ke Excel: " + e.getMessage())
} finally {
	if (fis != null) fis.close()
	if (wb != null) wb.close()
}

WebUI.comment('✅ Appended to Excel: ' + filePath)

/* =========================
 * SIGN OUT
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))
WebUI.waitForPageLoad(20)
WebUI.closeBrowser()
