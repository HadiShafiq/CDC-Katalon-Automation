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

	// Ensure datepicker is visible
	TestObject dp = new TestObject('dp')
	dp.addProperty("xpath", ConditionType.EQUALS, "//*[@id='ui-datepicker-div']")
	WebUI.waitForElementVisible(dp, 20)

	def parts = yyyyMmDd.split('-')
	int targetYear  = parts[0] as int
	int targetMonth = parts[1] as int   // 1..12
	String targetDay = String.valueOf(parts[2] as int) // "01" -> "1"

	// English month names used by jQuery UI datepicker
	Map<String, Integer> monthMap = [
		"January":1,"February":2,"March":3,"April":4,"May":5,"June":6,
		"July":7,"August":8,"September":9,"October":10,"November":11,"December":12
	]

	TestObject monthObj = new TestObject('dpMonth')
	monthObj.addProperty("xpath", ConditionType.EQUALS, "//*[@id='ui-datepicker-div']//span[@class='ui-datepicker-month']")

	TestObject yearObj = new TestObject('dpYear')
	yearObj.addProperty("xpath", ConditionType.EQUALS, "//*[@id='ui-datepicker-div']//span[@class='ui-datepicker-year']")

	TestObject nextBtn = new TestObject('dpNext')
	nextBtn.addProperty("xpath", ConditionType.EQUALS, "//*[@id='ui-datepicker-div']//a[contains(@class,'ui-datepicker-next')]")

	TestObject prevBtn = new TestObject('dpPrev')
	prevBtn.addProperty("xpath", ConditionType.EQUALS, "//*[@id='ui-datepicker-div']//a[contains(@class,'ui-datepicker-prev')]")

	// Navigate month/year until correct (safety max 48 clicks)
	int guard = 0
	while (guard < 48) {
		String curMonthName = WebUI.getText(monthObj).trim()
		int curMonth = monthMap.get(curMonthName)
		int curYear = WebUI.getText(yearObj).trim() as int

		if (curYear == targetYear && curMonth == targetMonth) break

		if (curYear < targetYear || (curYear == targetYear && curMonth < targetMonth)) {
			WebUI.click(nextBtn)
		} else {
			WebUI.click(prevBtn)
		}
		WebUI.delay(1) // must be number, not string
		guard++
	}

	// Click the day (avoid other-month/disabled cells)
	String dayXpath =
		"//*[@id='ui-datepicker-div']//td[" +
		"not(contains(@class,'ui-datepicker-other-month')) and " +
		"not(contains(@class,'ui-state-disabled'))" +
		"]//a[normalize-space(.)='${targetDay}']"

	TestObject dayObj = new TestObject("day_" + targetDay)
	dayObj.addProperty("xpath", ConditionType.EQUALS, dayXpath)

	WebUI.waitForElementClickable(dayObj, 20)
	WebUI.click(dayObj)
}

/* =========================================================
 * 8) BROWSER SETUP
 * ========================================================= */

/* PATH HADI */
String chromeBinary = "C:\\Users\\hadishafiq\\Downloads\\chrome-win64\\chrome-win64\\chrome.exe"
String chromeDriverPath = "C:\\Users\\hadishafiq\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe"

/* PATH ATIKAH
String chromeBinary = "C:\\Users\\nurul.atikah\\Documents\\CDC - Work\\Automation\\Automation Testing Browser FIles\\chrome-win64\\chrome-win64\\chrome.exe"
String chromeDriverPath = "C:\\Users\\nurul.atikah\\Documents\\CDC - Work\\Automation\\Automation Testing Browser FIles\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe"*/


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

/* =========================
 * DLOA - Simple Quote Respond
 * ========================= */

WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Click Simple Quote Respond'))

String targetSQ = sqNo
boolean found = false

for (int pageIndex = 1; pageIndex <= 10; pageIndex++) {

	WebUI.comment("Checking page index: " + pageIndex + " for SQ No: " + targetSQ)

	TestObject targetRow = new TestObject("targetRow_" + pageIndex)
	targetRow.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId_data']//tr[contains(.,'${targetSQ}')]"
	)

	if (WebUI.verifyElementPresent(targetRow, 3, FailureHandling.OPTIONAL)) {
		WebUI.comment("✅ Found SQ Number: " + targetSQ + " on page index " + pageIndex)

		TestObject procurementTitle = new TestObject("procurementTitle_" + pageIndex)
		procurementTitle.addProperty("xpath", ConditionType.EQUALS,
			"(//*[@id='_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId_data']//tr[contains(.,'${targetSQ}')]//*[starts-with(@id,'_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId:')]/span)[1]"
		)

		c(procurementTitle, 20)
		waitBlockUI(20)

		found = true
		break
	}

	if (pageIndex < 10) {
		TestObject nextPage = new TestObject("page_" + (pageIndex + 1))
		nextPage.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId_paginator_bottom']/span[4]/span[${pageIndex + 1}]"
		)

		if (WebUI.verifyElementClickable(nextPage, FailureHandling.OPTIONAL)) {
			WebUI.comment("SQ not found on page " + pageIndex + ". Moving to page " + (pageIndex + 1))
			c(nextPage, 20)
			waitBlockUI(20)
			WebUI.delay(1)
		}
	}
}

assert found : "❌ SQ Number not found from pagination 1 until 10: " + targetSQ

//WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Click Procument Tittle'))

int loopCount = 2

for (int i = 0; i < loopCount; i++) {

    String xpath = "(//input[contains(@id,'questionTbl:') and contains(@id,':ratePerUom')])[" + (i + 1) + "]"

    TestObject unitPriceField = new TestObject("unitPriceField_" + i)
    unitPriceField.addProperty("xpath", ConditionType.EQUALS, xpath)

    WebUI.comment("Fill Unit Price row #" + (i + 1))

    t(unitPriceField, UnitPrice, 20)
    waitBlockUI(20)
    WebUI.delay(1)

    WebUI.sendKeys(unitPriceField, Keys.chord(Keys.TAB))
    waitBlockUI(20)
    WebUI.delay(1)
}

WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Submit Respond Simple Quote'))
waitBlockUI(20)
WebUI.delay(3)

WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Click Soft Cert Sign Button'))
waitBlockUI(20)
WebUI.delay(3)

/* =========================
 * Sign Out
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))

// wait until logout is completed (choose one)
WebUI.waitForPageLoad(20)
WebUI.closeBrowser()


