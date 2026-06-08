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

/**
 * For Bumiputera or non Bumi
 * Yes = Bumiputera
 * No = non Bumiputera
 */

def clickBumiputera(String option) {

	String boxXpath

	if (option.equalsIgnoreCase("Yes")) {
		boxXpath = "//table[contains(@id,'bumiId')]//label[normalize-space()='Yes']/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
	} else if (option.equalsIgnoreCase("No")) {
		boxXpath = "//table[contains(@id,'bumiId')]//label[normalize-space()='No']/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
	} else {
		throw new Exception("Invalid option: " + option)
	}

	TestObject obj = new TestObject("bumi_" + option)
	obj.addProperty("xpath", ConditionType.EQUALS, boxXpath)

	WebUI.waitForElementClickable(obj, 10)
	WebUI.scrollToElement(obj, 5)

	WebElement element = WebUiCommonHelper.findWebElement(obj, 10)
	WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(element))

	waitBlockUI(10)
	WebUI.delay(0.3)
}
/**
 * For Muslim or non Muslim
 * Yes = Muslim
 * No = non Muslim
 */
def clickReligion(int religion) {

	String boxXpath

	switch (religion) {
		case 1: // Muslim
			boxXpath = "//table[contains(@id,'religionId')]//label[normalize-space()='Muslim']/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
			break

		case 2: // Non - Muslim
			boxXpath = "//table[contains(@id,'religionId')]//label[contains(normalize-space(),'Non')]/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
			break

		default:
			throw new Exception("Invalid religion value: ${religion}. Use 1=Muslim, 2=Non Muslim.")
	}

	TestObject obj = new TestObject("religion_" + religion)
	obj.addProperty("xpath", ConditionType.EQUALS, boxXpath)

	WebUI.waitForElementClickable(obj, 10)
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

//Activation finder helper
// Helper function
def findByXPath(String xpath) {
	TestObject obj = new TestObject()
	obj.addProperty('xpath', ConditionType.EQUALS, xpath)
	return obj
}

/* =========================
 * OPEN APPLICATION
 * Purpose:
 * - open outlook email
 * - maximize browser
 * - wait initial page load
 * ========================= */
WebUI.navigateToUrl('https://yopmail.com/en/')
WebUI.maximizeWindow()
waitBlockUI(20)

//enter email + next
t(findTestObject('Object Repository/EmailVerification/Yopmail/EmailInput'), SupplierEmail)
c(findTestObject('Object Repository/EmailVerification/Yopmail/HomeNext'))
waitBlockUI(20)

// Switch into Yopmail email iframe first
WebUI.switchToFrame(findByXPath("//iframe[@id='ifmail']"), 10)

// Activation Key grabber
WebElement element = WebUI.findWebElement(
    findByXPath("//li[contains(text(),'Activation Key')]"), 10
)

// Get full text and extract the key
String fullText = element.getText()
String activationKey = fullText.split(': ')[1].trim()
println("Activation Key: " + activationKey)

	
	//Clink Activate link to activate account
	c(findTestObject('Object Repository/EmailVerification/EmailPage/AccountActivationLink'))
	waitBlockUI(10)
	
		// Get all window handles
		List<String> tabs = driver.getWindowHandles().toList()
		
		// Switch to new tab (index 1 = second tab)
		WebUI.switchToWindowIndex(1)
		waitBlockUI(5)


//Input activation key and input to 'kod Pengaktifan'
WebUI.waitForElementPresent(findTestObject('Object Repository/EmailVerification/PengaktifanPenggunaPembekal/KodPengaktifan'),10)
t(findTestObject('Object Repository/EmailVerification/PengaktifanPenggunaPembekal/KodPengaktifan'),activationKey)
	waitBlockUI(10)
//ID Log Masuk
t(findTestObject('Object Repository/EmailVerification/PengaktifanPenggunaPembekal/IDLogMasukBox'),IDLogMasuk)
//Next
c(findTestObject('Object Repository/EmailVerification/PengaktifanPenggunaPembekal/Serah'))
waitBlockUI(5)

c(findTestObject('Object Repository/EmailVerification/PengaktifanPenggunaPembekal/Ya'))
waitBlockUI(5)

//No eP grabber

	// Grab the success message element
	WebElement ePElement = WebUI.findWebElement(
		findByXPath("//span[contains(@class,'ui-messages-info-detail')]"), 10
	)
	
	// Get full text
	String ePFullText = ePElement.getText()
	// "Anda telah berjaya mengaktifkan Akaun Asas anda dengan No. eP berikut: eP-0101A0004 . Sila semak emel anda..."
	
	// Extract eP number - split by 'berikut: ' and take what's after, then trim the space and dot
	String ePNumber = ePFullText.split('berikut: ')[1].split(' \\.')[0].trim()
	println("eP Number: " + ePNumber)  
	
	// Store as Global Variable if needed across test cases
	GlobalVariable.ePNumber = ePNumber
	

	
	