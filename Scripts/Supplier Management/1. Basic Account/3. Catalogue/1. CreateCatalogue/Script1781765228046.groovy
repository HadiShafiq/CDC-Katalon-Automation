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
// Helper function - must be in every script that uses findByXPath
def findByXPath(String xpath) {
	TestObject obj = new TestObject()
	obj.addProperty('xpath', ConditionType.EQUALS, xpath)
	return obj
}

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
 * Set unit price / rate by popup row index.
 * Used for popup table:
 * specAnswerTbl:{rowIndex}:ratePerUomAns
 */
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


/* =========================================================
 * 6) SPECIAL RADIO / OPTION HELPER
 * ========================================================= */

/**
 * Click Account type button by option number.
 * 1 = Basic
 * 2 = MOF
 * 3 = G2G
 */
def clickAccountType(int option) {

	String xpath =
		"(//div[contains(@class,'ui-radiobutton')]//input[@type='radio'])[${option}]"

	TestObject obj = new TestObject("accountType_" + option)
	obj.addProperty("xpath", ConditionType.EQUALS, xpath)

	WebUI.waitForElementPresent(obj, 10)
	WebUI.scrollToElement(obj, 5)

	WebElement element = WebUiCommonHelper.findWebElement(obj, 10)

	WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(element))

	waitBlockUI(10)
	WebUI.delay(0.3)
}


/* =========================================================
 * 7) CALENDAR PICKER DATE - Supplier Registration Page only
 * ========================================================= */

def pickDate(String yyyyMmDd) {

	def parts = yyyyMmDd.split('-')
	int targetYear  = parts[0] as int
	int targetMonth = parts[1] as int
	String targetDay = (parts[2] as int).toString()

	TestObject monthObj = new TestObject()
	monthObj.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='ui-datepicker-div']//select[contains(@class,'ui-datepicker-month')]")

	TestObject yearObj = new TestObject()
	yearObj.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='ui-datepicker-div']//select[contains(@class,'ui-datepicker-year')]")

	WebUI.waitForElementVisible(monthObj, 10)

	WebUI.selectOptionByValue(monthObj, String.valueOf(targetMonth - 1), false)
	WebUI.selectOptionByValue(yearObj, String.valueOf(targetYear), false)

	String dayXpath =
		"//*[@id='ui-datepicker-div']//td[not(contains(@class,'ui-datepicker-other-month')) " +
		"and not(contains(@class,'ui-state-disabled'))]//a[normalize-space()='" + targetDay + "']"

	TestObject dayObj = new TestObject()
	dayObj.addProperty("xpath", ConditionType.EQUALS, dayXpath)

	WebUI.waitForElementClickable(dayObj, 10)
	WebUI.click(dayObj)
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
wVisible(findTestObject('Object Repository/Home Page/LanguageSelect'), 20)
WebUI.selectOptionByValue(findTestObject('Object Repository/Home Page/LanguageSelect'), 'en_US', true)
	waitBlockUI(20)
	WebUI.delay(0.5)
	WebUI.delay(1)
	
//Login
	c(findTestObject('Object Repository/Home Page/LoginHomePage'))
	t(findTestObject('Object Repository/Home Page/LoginPage/LoginID'), Username)
	t(findTestObject('Object Repository/Home Page/LoginPage/Password'), Password)
	c(findTestObject('Object Repository/Home Page/LoginPage/LoginButtonEng'))
	
//Homepage 
	//Upload Catalogue
	c(findTestObject('Object Repository/Home Page/SupplierPage/UploadCatalogue'))
	
	//Create New Catalogue Item page
	t(findTestObject('Object Repository/CreateCatalogue/CreateNewCatalogueItemPage/ItemCode'),ItemCode)
	//Search
	c(findTestObject('Object Repository/CreateCatalogue/CreateNewCatalogueItemPage/CatalogueSearch'))
	//Request Item
	c(findTestObject('Object Repository/CreateCatalogue/CreateNewCatalogueItemPage/CatalogueRequestItem'))
	
	//Item Code Request Form
	
	//Item type
	selectDropdownByIndex(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/ItemTypeDropdown'), ItemType)
	
	if (ItemType == 0) {		//Product
		// Menu for Product
		
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/ItemNameForm'), ItemName)
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/RequestDescriptionForm'), RequestDescription)
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/TypeForm'),Type)
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/Brand'), ItemBrand)
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/MeasurementForm'),Measurement)
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/ColorForm'),Color)
		c(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/ItemCodeRequestSubmit'))
		
				/// Grab the success message element
WebElement rncElement = WebUI.findWebElement(
    findByXPath("//span[contains(@class,'ui-messages-info-detail') and contains(text(),'has been successfully submitted')]"), 10
)

// Get full text
String rncFullText = rncElement.getText()

// Extract RNC number
String rncNumber = rncFullText.split('request ')[1].split(' has')[0].trim()
println("RNC Number: " + rncNumber)  
		} 
	
		
		
	 else if (ItemType == 1 ) {  //Service
		// Menu for Service
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/ItemNameForm'), ItemName)
		t(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/RequestDescriptionForm'), RequestDescription)
		c(findTestObject('Object Repository/CreateCatalogue/ItemCodeRequestForm/ItemCodeRequestSubmit'))
	
		/// Grab the success message element
WebElement rncElement = WebUI.findWebElement(
    findByXPath("//span[contains(@class,'ui-messages-info-detail') and contains(text(),'has been successfully submitted')]"), 10
)

// Get full text
String rncFullText = rncElement.getText()
} 
/* Extract RNC number 
String rncNumber = rncFullText.split('request ')[1].split(' has')[0].trim()
println("RNC Number: " + rncNumber)  
		
		
		GlobalVariable.rncNumber = rncNumber */

//Excel appender
/* =========================
 * SUCCESS MESSAGE
 * Purpose:
 * - wait for loader disappear
 * - capture success message
 * - extract dynamic RNC number
 * ========================= */
TestObject blockUI = new TestObject('blockUI')
blockUI.addProperty("xpath", ConditionType.EQUALS,
	"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
)

if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
	WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
}

TestObject msgObj = new TestObject('msg_RNC_saved')
msgObj.addProperty("xpath", ConditionType.EQUALS,
	"//span[contains(@class,'ui-messages-info-detail') and contains(.,'has been successfully submitted')]"
)

WebUI.waitForElementVisible(msgObj, 30)

String msg = ""
for (int i = 0; i < 15; i++) {
	msg = WebUI.getText(msgObj, FailureHandling.OPTIONAL)
	if (msg != null && msg.contains("RNC")) break
	WebUI.delay(1)
}

msg = (msg == null) ? "" : msg.trim()
WebUI.comment("Message: " + msg)

def matcher = (msg =~ /(RNC\d+\-\d+)/)
String rncNumber = matcher.find() ? matcher.group(1) : ""

if (rncNumber == "") {
	WebUI.takeScreenshot()
	assert false : "❌ RNC number not found. Message was: " + msg
}
WebUI.comment("✅ Captured RNC No: " + rncNumber)

/* =========================
 * EXCEL APPEND
 * Purpose:
 * - append RNC number and message into Excel file
 * ========================= */
String baseDir = System.getProperty("user.home") + "/Desktop/PrepDataFileNumber"
new File(baseDir).mkdirs()
String filePath = baseDir + "/BasicAccount_RNC_Numbers.xlsx"
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
	header.createCell(1).setCellValue("RNC No")
	header.createCell(2).setCellValue("Message")
}

if (fis != null) fis.close()

int nextRow = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : sheet.getLastRowNum() + 1
def row = sheet.createRow(nextRow)

row.createCell(0).setCellValue(now)
row.createCell(1).setCellValue(rncNumber)
row.createCell(2).setCellValue(msg)

FileOutputStream fos = new FileOutputStream(filePath)
wb.write(fos)
fos.close()
wb.close()

WebUI.comment("✅ Appended to Excel: " + filePath)