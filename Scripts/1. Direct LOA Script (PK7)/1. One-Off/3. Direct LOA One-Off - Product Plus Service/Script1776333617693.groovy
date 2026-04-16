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

/* =========================
 * HELPERS
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
 * BROWSER SETUP
 * Purpose:
 * - launch Chrome in clean guest/incognito mode
 * - disable password manager prompts
 * ========================= */

String chromeBinary = "C:\\Users\\hadishafiq\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe"
String chromeDriverPath = "C:\\Users\\hadishafiq\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe"

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
 * CREATE LOA
 * Purpose:
 * - open Direct LOA creation screen
 * ========================= */
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Dashboard/Create LOA'), 20)
waitBlockUI(30)

/* =========================
 * SUPPLIER POPUP
 * Purpose:
 * - open supplier search popup
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Button Search Supplier - Direct LOA'), 20)
waitBlockUI(30)

/* =========================
 * SUPPLIER SEARCH
 * Purpose:
 * - select supplier dropdown filter
 * - input supplier name
 * - search supplier
 * ========================= */
selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/LOA - Supplier Dropdown'), 1)

t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Key In Business Name'), SupplierName, 20)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Button Search Supplier'), 20)
waitBlockUI(30)

/* =========================
 * SUPPLIER SELECTION
 * Purpose:
 * - choose supplier row from result
 * - confirm selection
 * ========================= */
TestObject supplierRow = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Click the Supplier')
wVisible(supplierRow, 20)
WebUI.scrollToElement(supplierRow, 2)

try {
	WebUI.click(supplierRow)
} catch (Exception e) {
	WebUI.enhancedClick(supplierRow, FailureHandling.OPTIONAL)
}
WebUI.delay(0.5)
WebUI.doubleClick(supplierRow, FailureHandling.OPTIONAL)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Select Supplier'), 20)
waitBlockUI(1)
WebUI.delay(2)

/* =========================
 * GENERAL INFORMATION - DROPDOWNS
 * Purpose:
 * - select procurement method and related dropdown values
 * ========================= */
TestObject DPM = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Procurement Method')

// first click may reset/close dropdown
WebUI.click(DPM)
WebUI.delay(1)

// second click opens properly
WebUI.click(DPM)
WebUI.delay(2)

// select procurement method
selectDropdownByIndex(DPM, ProcurementMethod)

selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/DropDown Procurement Type Category'), ProcurementTypeCategory)
waitBlockUI(20)
WebUI.delay(0.5)
selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Quotation_Tender Type'), QuotationTenderType)
waitBlockUI(20)
WebUI.delay(0.5)
selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Reason PK 7'), Reason)
waitBlockUI(20)
WebUI.delay(0.5)
selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Procurement Category'), ProcurementCategory)
waitBlockUI(20)
WebUI.delay(0.5)
/* =========================
 * GENERAL INFORMATION - INPUTS
 * Purpose:
 * - input title, reference no, LOA price
 * - trigger blur/calculation with TAB
 * ========================= */
TestObject ministryTA = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Title Letter of Acceptance')
t(ministryTA, TitleLetterofAcceptance, 20)
WebUI.sendKeys(ministryTA, Keys.chord(Keys.TAB))
WebUI.delay(1)


/* =========================
 * File Reference
 * ========================= */
TestObject safeArea = new TestObject('safeArea')
safeArea.addProperty(
    "xpath",
    ConditionType.EQUALS,
    "//label[normalize-space()='File Reference No.']"
)

TestObject fileRef1 = findTestObject(
    'Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/File Reference No_1'
)

String ref1Input = FileReference1.toString().trim()

WebUI.waitForElementVisible(fileRef1, 20)
WebUI.waitForElementClickable(fileRef1, 20)
WebUI.click(fileRef1)
WebUI.delay(0.5)

WebUI.sendKeys(fileRef1, Keys.chord(Keys.CONTROL, 'a'))
WebUI.delay(0.3)
WebUI.sendKeys(fileRef1, Keys.chord(Keys.BACK_SPACE))
WebUI.delay(0.5)

for (char ch : ref1Input.toCharArray()) {
    WebUI.sendKeys(fileRef1, ch.toString())
    WebUI.delay(0.15)
}

WebUI.click(safeArea, FailureHandling.OPTIONAL)

waitBlockUI(10)
WebUI.delay(1)

String finalRef1 = WebUI.getAttribute(fileRef1, 'value')
println("Final File Reference No_1 = " + finalRef1)

t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/File Reference No_2'), FileReference2, 20)
WebUI.delay(1)

/* =========================
 * LOA Price
 * ========================= */

TestObject loaSafeArea = new TestObject('loaSafeArea')
loaSafeArea.addProperty(
    "xpath",
    ConditionType.EQUALS,
    "//label[contains(normalize-space(),'LOA Offered Price')]"
)
TestObject loaPrice = findTestObject(
    'Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/LOA Offered Price (RM)'
)

String rawInput = LOAOfferedPrice.toString().replace(",", "").trim()

WebUI.waitForElementVisible(loaPrice, 20)
WebUI.waitForElementClickable(loaPrice, 20)
WebUI.click(loaPrice)
WebUI.delay(0.5)

WebUI.sendKeys(loaPrice, Keys.chord(Keys.CONTROL, 'a'))
WebUI.delay(0.3)
WebUI.sendKeys(loaPrice, Keys.chord(Keys.BACK_SPACE))
WebUI.delay(0.5)

for (char ch : rawInput.toCharArray()) {
    WebUI.sendKeys(loaPrice, ch.toString())
    WebUI.delay(0.2)
}

// click label instead of body
WebUI.click(loaSafeArea, FailureHandling.OPTIONAL)


waitBlockUI(10)
WebUI.delay(1)

String finalValue1 = WebUI.getAttribute(loaPrice, 'value')
println("Final LOA Offered Price = " + finalValue1)

/* =========================
 * Service Tax
 * ========================= */

TestObject taxSafeArea = new TestObject('taxSafeArea')
taxSafeArea.addProperty(
    "xpath",
    ConditionType.EQUALS,
    "//label[contains(normalize-space(),'Sales Tax / Service Tax (RM)')]"
)

TestObject serviceTax = findTestObject(
    'Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Service Tax'
)

String rawInput1 = SalesTaxValue.toString().replace(",", "").trim()

WebUI.waitForElementVisible(serviceTax, 20)
WebUI.waitForElementClickable(serviceTax, 20)
WebUI.click(serviceTax)
WebUI.delay(0.5)

WebUI.sendKeys(serviceTax, Keys.chord(Keys.CONTROL, 'a'))
WebUI.delay(0.3)
WebUI.sendKeys(serviceTax, Keys.chord(Keys.BACK_SPACE))
WebUI.delay(0.5)

for (char ch : rawInput1.toCharArray()) {
    WebUI.sendKeys(serviceTax, ch.toString())
    WebUI.delay(0.2)
}

WebUI.click(taxSafeArea, FailureHandling.OPTIONAL)

waitBlockUI(10)
WebUI.delay(1)

String finalValue3 = WebUI.getAttribute(serviceTax, 'value')
println("Final Sales Tax / Service Tax (RM) = " + finalValue3)

/* =========================
 * CONTRACT DETAILS
 * Purpose:
 * - select fulfilment type, performance bond, contract type
 * - input duration
 * ========================= */
selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Fullfilment Type'), FulfilmentType)
WebUI.delay(1)

selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Performance Bond'), PerformanceBond)
WebUI.delay(5)

//Verification Radio Button
// value: 1 = Yes, 2 = No
def clickRequiredOnlineVerification(def value) {

	if (value == null || value.toString().trim() == "") {
		assert false : "❌ Required Online Verification value is empty"
	}

	int intValue = value.toString().trim().toInteger()
	String label = (intValue == 1) ? "Yes" : "No"

	TestObject opt = new TestObject("requiredOnlineVerification_" + label)
	opt.addProperty("xpath", ConditionType.EQUALS,
		"//label[contains(normalize-space(.),'Required Online Verification')]/following::label[normalize-space(.)='${label}'][1]"
	)

	WebUI.waitForElementClickable(opt, 20)
	WebUI.click(opt)
}

// call your variable here
clickRequiredOnlineVerification(RequiredOnlineVerification)

selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Dropdown Contract Type'), ContractType)

// Month Duration
TestObject durationObj = findTestObject(
	'Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Duration'
)
WebUI.waitForElementVisible(durationObj, 20)
WebElement durationEl = WebUiCommonHelper.findWebElement(durationObj, 20)

WebUI.executeJavaScript(
    """
    arguments[0].value = arguments[1];
    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
	arguments[0].blur();
    """,
    Arrays.asList(durationEl, ContractPeriod)
)

waitBlockUI(20)
WebUI.delay(1)

String finalValue2 = WebUI.getAttribute(durationObj, 'value')
println("Final Duration = " + finalValue2)

/* =========================
 * DATE PICKER
 * Purpose:
 * - open and select date
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Date Picker icon'), 20)
WebUI.delay(1)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Click Date'), 20)
waitBlockUI(20)

/* =========================
 * REQUIRED AGREEMENT
 * Purpose:
 * - choose agreement value from dropdown
 * - if No selected, handle confirmation popup
 * ========================= */
import java.util.Arrays

def clickAgreementReq(def value) {

	WebUI.comment("DEBUG Agreement raw value = [" + value + "]")
	WebUI.comment("DEBUG Agreement class = " + (value == null ? "null" : value.getClass().getName()))

	if (value == null || value.toString().trim() == "") {
		assert false : "❌ Agreement value is empty"
	}

	String rawValue = value.toString().trim()
	int intValue = rawValue.toInteger()

	WebUI.comment("DEBUG Agreement intValue = " + intValue)

	TestObject dd = new TestObject("agreementReq_dropdown")
	dd.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='_scCreateManualSourcing_WAR_NGePportlet_:form:agreementReqPk7Per']/div[3]/span"
	)

	WebUI.waitForElementClickable(dd, 20)
	WebUI.click(dd)

	String optXpath = (intValue == 1)
		? "//*[@id='_scCreateManualSourcing_WAR_NGePportlet_:form:agreementReqPk7Per_panel']//li[2]"
		: "//*[@id='_scCreateManualSourcing_WAR_NGePportlet_:form:agreementReqPk7Per_panel']//li[3]"

	TestObject opt = new TestObject("agreementReq_option")
	opt.addProperty("xpath", ConditionType.EQUALS, optXpath)

	WebUI.waitForElementClickable(opt, 20)
	WebUI.click(opt)

	if (intValue == 2) {
		WebUI.comment("DEBUG value = 2, clicking popup Yes")

		TestObject popYes = new TestObject('popYes')
		popYes.addProperty("xpath", ConditionType.EQUALS,
			"//button[contains(@onclick,'qtChangeAgreementReqDlg.hide') and contains(@onclick,'agreementReqPG') and .//span[normalize-space()='Yes']]"
		)

		WebUI.waitForElementPresent(popYes, 20)
		WebUI.waitForElementVisible(popYes, 20)
		WebUI.scrollToElement(popYes, 2)

		def el = WebUI.findWebElement(popYes, 20)
		WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(el))
	}
}

clickAgreementReq(Agreement)

/* =========================
 * CATEGORY CODE
 * Purpose:
 * - open category code section
 * - tick category checkbox
 * - confirm and save selection
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Category Code'), 20)
waitBlockUI(30)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Category Code/Click CheckBox Category Code'), 20)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Category Code/Click OK Category Code'), 20)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Category Code/Click Yes Pop Up Category Code'), 20)
waitBlockUI(30)

/* =========================
 * LOA & ATTACHMENT
 * Purpose:
 * - open LOA attachment section
 * - set LOA date
 * - open signer popup
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Letter Of Acceptance And Attachment'), 20)
waitBlockUI(30)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Date Picker LOA and Attachment'), 20)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/General infomation Tab/Click Date'), 20)
waitBlockUI(20)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Button LOA Signer'), 20)
WebUI.delay(1)
waitBlockUI(10)

/* =========================
 * SIGNER SEARCH
 * Purpose:
 * - filter by username
 * - search LOA signer
 * - select signer row
 * ========================= */
selectDropdownByIndex(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Dropdown Username'), 1)
WebUI.delay(0.5)

t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Input User Name'), LOASigner, 20)
WebUI.delay(0.5)
waitBlockUI(30)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Search User Name'), 20)
WebUI.delay(0.5)
waitBlockUI(30)

TestObject userRow = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Click in user Table')
wVisible(userRow, 20)
WebUI.scrollToElement(userRow, 2)

try {
	WebUI.click(userRow)
} catch (Exception e) {
	WebUI.enhancedClick(userRow, FailureHandling.OPTIONAL)
}
WebUI.delay(0.5)
WebUI.doubleClick(userRow, FailureHandling.OPTIONAL)

/* =========================
 * DOCUMENT UPLOAD
 * Purpose:
 * - upload LOA signer document
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Click upload Icon'), 3)
up(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Choose File'),
	'C:\\Users\\hadishafiq\\Desktop\\File\\File.pdf',3)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Upload Icon LOA Signer Document'), 20)
waitBlockUI(30)

/* =========================
 * ZONE ITEM - Zonal
 * Purpose:
 * To choose radio button for zonal
 *
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Zone Item'), 20)


def clickZoneLocRadio(int option) {
	String xpath

	switch(option) {
		case 1:
			// Yes
			xpath = "//*[@id='_scCreateManualSourcing_WAR_NGePportlet_:form:zoneLocFlg']/tbody/tr/td[1]/div/div[2]"
			break

		case 2:
			// No
			xpath = "//*[@id='_scCreateManualSourcing_WAR_NGePportlet_:form:zoneLocFlg']/tbody/tr/td[3]/div/div[2]"
			break

		default:
			throw new Exception("Invalid option. Use 1 for Yes or 2 for No.")
	}

	TestObject obj = new TestObject("zoneLocRadio")
	obj.addProperty("xpath", ConditionType.EQUALS, xpath)

	c(obj, 20)
}

// 1 = Yes
// 2 = No
clickZoneLocRadio(ZoneLocation)
waitBlockUI(30)
WebUI.delay(1)

if (ZoneLocation.toString().trim() == "1") {

	// Zonal Coverage DropDown
	selectDropdownByIndex(
		findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/Zone Item/Zonal/Zonal Coverage DropDown'),
		ZonalCoverage
	)
	waitBlockUI(30)
	WebUI.delay(0.5)

	def zoneGroups = [
		"Zone A": [1, 2, 3, 4, 5, 6, 7, 8, 9],
		"Zone B": [10, 11, 12, 13, 14, 15, 16]
	]

	def tickZoneTreeByIndex = { int index ->
		String xpath = "//*[@id='_scCreateManualSourcing_WAR_NGePportlet_:form:treeZoneGeneralPopup:${index}']//div[contains(@class,'ui-chkbox-box')]"

		TestObject obj = new TestObject("zoneTreeTick_" + index)
		obj.addProperty("xpath", ConditionType.EQUALS, xpath)

		WebUI.waitForElementVisible(obj, 20)
		WebUI.waitForElementClickable(obj, 20)
		WebUI.scrollToElement(obj, 20)
		WebUI.click(obj)

		waitBlockUI(30)
		WebUI.delay(0.5)
	}

	// MAIN LOOP
	zoneGroups.each { zoneName, indexes ->

		WebUI.comment("Processing " + zoneName)

		// open popup
		c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/Zone Item/Zonal/Add Button Zone'))
		waitBlockUI(30)
		WebUI.delay(1)

		// set zone name
		TestObject zoneNameObj = findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/Zone Item/Zonal/Zone Name')
		WebElement zoneNameEl = WebUiCommonHelper.findWebElement(zoneNameObj, 20)

		WebUI.executeJavaScript(
			"arguments[0].value = arguments[1];",
			Arrays.asList(zoneNameEl, zoneName)
		)

		WebUI.delay(1)

		// tick all indexes for this zone
		indexes.each { idx ->
			tickZoneTreeByIndex(idx)
		}

		// click add locality
		c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/Zone Item/Zonal/Add Locality'))
		waitBlockUI(30)
		WebUI.delay(1)
	}

} else if (ZoneLocation.toString().trim() == "2") {

	WebUI.comment("Zone Location = No. Skip zonal coverage and zone locality section.")

}
/* =========================
 * ZONE ITEM - PRODUCT (LOOPING)
 * Purpose:
 * - open product section
 * - loop add product details
 * ========================= */

// Click Side Menu Zone Item
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Zone Item'), 20)
waitBlockUI(30)
WebUI.delay(1)

// Set number of loops (1-10)
int loopCount = 2  // Set how many times you want the loop to run (1-10)

for (int i = 1; i <= loopCount; i++) {
	WebUI.comment("Loop #${i} of ${loopCount}")

	// Click Action Item for Product on every loop, including first loop
		c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Click Action Item for Product'), 20)
		waitBlockUI(20)
		WebUI.delay(1)
	

	// Input Specification for Product 1 (Dynamic Text + Loop Index)
	TestObject spec1Product = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input Specification 1 TextBox')
	WebUI.waitForElementVisible(spec1Product, 20)
	WebUI.waitForElementClickable(spec1Product, 20)
	WebUI.click(spec1Product)
	WebUI.clearText(spec1Product)

	String productSpec1WithLoopIndex = ProductSpecification1 + i
	WebUI.setText(spec1Product, productSpec1WithLoopIndex)

	waitBlockUI(30)
	WebUI.delay(2)

	// Input UOM for Product
	TestObject uom = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input UOM')
	wVisible(uom, 20)
	WebUI.click(uom)
	t(uom, ProductUOM, 20)
	WebUI.delay(1)
	WebUI.sendKeys(uom, Keys.chord(Keys.ENTER))
	waitBlockUI(30)
	WebUI.delay(1)

	// Price Type (REAL <select>) - 0-based data
	wVisible(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Dropdown Price Type'), 20)
	WebUI.selectOptionByIndex(
		findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Dropdown Price Type'),
		toInt(PriceType)
	)

	waitBlockUI(30)
	WebUI.delay(1)

	// Input Unit Price
	t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Unit Price(RM)'),
		ProductUnitPrice, 20
	)
	waitBlockUI(30)
	WebUI.delay(1)

	// Input Quantity
	TestObject ProductQty = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Quantity')
	t(ProductQty, ProductQuaty, 20)
	WebUI.sendKeys(ProductQty, Keys.chord(Keys.TAB))
	waitBlockUI(30)
	WebUI.delay(1)

	/* =========================
	 * ADDITIONAL SPECIFICATION
	 * ========================= */

	// Click Action Item for Specification
	TestObject specBtn = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Action Item Click Specification')
	WebUI.waitForElementVisible(specBtn, 20)
	WebUI.waitForElementClickable(specBtn, 20)
	WebUI.click(specBtn)

	TestObject clickSpec = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Click Specification')
	WebUI.waitForElementClickable(clickSpec, 20)
	WebUI.click(clickSpec)

	waitBlockUI(20)
	WebUI.delay(2)

	// Input Specification 2 TextBox (Dynamic + Loop Index)
	TestObject spec2Product = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input Specification 2 TextBox')

	String productSpec2WithLoopIndex = ProductSpecification2 + i
	WebUI.setText(spec2Product, productSpec2WithLoopIndex)

	waitBlockUI(30)
	WebUI.delay(3)
}

/* =========================
 * ZONE ITEM - SERVICES (LOOPING)
 * Purpose:
 * - open services section
 * - loop add services details
 * ========================= */
// Set number of loops (1-10)
int loopCountService = 2  // Set how many times you want the loop to run (1-10)

for (int i = 1; i <= loopCountService; i++) {
	WebUI.comment("Loop #${i} of ${loopCountService}")

		// Click Action Item for Service after the first loop starts
		c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Click Action Item for Service'), 20)
		waitBlockUI(20)
		WebUI.delay(1)

	// Input Specification for Service 1 (Dynamic Text + Loop Index)
	TestObject spec1Service = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Input Specification 1 TextBox')
	WebUI.waitForElementVisible(spec1Service, 20)
	WebUI.waitForElementClickable(spec1Service, 20)
	WebUI.click(spec1Service)                 // focus
	WebUI.clearText(spec1Service)             // clear existing value
	String serviceSpec1WithLoopIndex = ServiceSpecification1 + i  // Add loop index to Specification Text
	WebUI.setText(spec1Service, serviceSpec1WithLoopIndex) // type new value
	waitBlockUI(30)
	WebUI.delay(1)

	// Input UOM for Service
	TestObject uomService = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Input UOM')
	wVisible(uomService, 20)
	WebUI.click(uomService)
	t(uomService, ServiceUOM, 20)
	WebUI.delay(2)
	WebUI.sendKeys(uomService, Keys.chord(Keys.ENTER))
	waitBlockUI(30)
	WebUI.delay(1)
	
	// Input Service Freq. per OM
	t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Freq per UOM'), FreqperUOM, 20)
	waitBlockUI(30)
	WebUI.delay(1)

	// Input Service Quantity
	t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Quantity'), ServiceQty, 20)
	waitBlockUI(30)
	WebUI.delay(1)
	
	// Input Service Unit Price(RM)
	t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Unit Price(RM)'), ServiceUnitPrice, 20)
	waitBlockUI(30)
	WebUI.delay(1)

	/* =========================
	 * ADDITIONAL SPECIFICATION
	 * ========================= */

	// Click Action Item for Specification (Service)
	TestObject specBtn = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Action Item Click Specification')
	WebUI.waitForElementVisible(specBtn, 20)
	WebUI.waitForElementClickable(specBtn, 20)
	WebUI.click(specBtn)
	
	TestObject clickSpec = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Click Specification Services')
	WebUI.waitForElementClickable(clickSpec, 20)
	WebUI.click(clickSpec)

	waitBlockUI(20)
	WebUI.delay(3)

	// Input Specification 2 TextBox for Service
	TestObject spec2Service = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Input Specification 2 TextBox')
	String serviceSpec2WithLoopIndex = ServiceSpecification2 + i  // Add loop index to Specification Text
	WebUI.setText(spec2Service, serviceSpec2WithLoopIndex) // type new value
	waitBlockUI(30)
	WebUI.delay(1)
}

/* =========================
 * SAVE / SUBMIT LOA
 * Purpose:
 * - navigate to payment deduction side menu
 * - submit LOA application
 * ========================= */
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Payment Deduction'), 20)
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Submit LOA Application'))

waitBlockUI(30)

/* =========================
 * SUCCESS MESSAGE
 * Purpose:
 * - wait for loader disappear
 * - capture success message
 * - extract dynamic LOA number
 * ========================= */
TestObject blockUI = new TestObject('blockUI')
blockUI.addProperty("xpath", ConditionType.EQUALS,
	"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
)

if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
	WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
}

TestObject msgObj = new TestObject('msg_LOA_saved')
msgObj.addProperty("xpath", ConditionType.EQUALS,
	"//span[contains(@class,'ui-messages-info-detail') and " +
	"contains(.,'Letter of Acceptance (LOA)') and contains(.,'is successfully submitted')]"
)

WebUI.waitForElementVisible(msgObj, 30)

String msg = ""
for (int i = 0; i < 15; i++) {
	msg = WebUI.getText(msgObj, FailureHandling.OPTIONAL)
	if (msg != null && msg.contains("LA")) break
	WebUI.delay(1)
}

msg = (msg == null) ? "" : msg.trim()
WebUI.comment("Message: " + msg)

def matcher = (msg =~ /(LA\d+)/)
String loaNo = matcher.find() ? matcher.group(1) : ""

if (loaNo == "") {
	WebUI.takeScreenshot()
	assert false : "❌ LOA number not found. Message was: " + msg
}
WebUI.comment("✅ Captured LOA No: " + loaNo)

/* =========================
 * EXCEL APPEND
 * Purpose:
 * - append LOA number and message into same Excel file
 * ========================= */
String filePath = "C:\\Users\\hadishafiq\\Desktop\\PrepData\\Direct_LOA_Non-Zonal_PK7_Product_AP_201_2026.xlsx"
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
	header.createCell(1).setCellValue("LOA No")
	header.createCell(2).setCellValue("Message")
}

if (fis != null) fis.close()

int nextRow = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : sheet.getLastRowNum() + 1
def row = sheet.createRow(nextRow)

row.createCell(0).setCellValue(now)
row.createCell(1).setCellValue(loaNo)
row.createCell(2).setCellValue(msg)

FileOutputStream fos = new FileOutputStream(filePath)
wb.write(fos)
fos.close()
wb.close()

WebUI.comment("✅ Appended to Excel: " + filePath)

/* =========================
 * SIGN OUT
 * Purpose:
 * - logout from system
 * - close browser
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))

WebUI.waitForPageLoad(20)
WebUI.closeBrowser()