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
 * ========================= */
WebUI.navigateToUrl('http://ngepsit.eperolehan.com.my/home')
WebUI.maximizeWindow()
waitBlockUI(20)

/* =========================
 * LANGUAGE
 * ========================= */
wVisible(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 20)
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)
waitBlockUI(20)
WebUI.delay(0.5)
WebUI.delay(1)

/* =========================
 * LOGIN
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
 * ========================= */
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

//Pending Delivery List Menu
c(findTestObject('Object Repository/DP - Add To Cart/Pending Delivery List/Click Contract List'))

//Dropdown Search Contract List: 
selectDropdownByIndex(findTestObject('Object Repository/DP - Add To Cart/Pending Delivery List/Dropdown Justification'), ddContractList)

//Input Contract List:
t(findTestObject('Object Repository/DP - Add To Cart/Pending Delivery List/Input Contract List'), ContractList)

//Click Button Search
c(findTestObject('Object Repository/DP - Add To Cart/Pending Delivery List/Button Search'))

/* =======================================================
 * XPATH FOR INPUT SCHEDULE & AS&WHEN
 * =======================================================*/
//SCHEDULE
TestObject scheduleIcon = new TestObject()
scheduleIcon.addProperty("xpath", ConditionType.EQUALS,
    "//div[contains(@class, 'ui-row-toggler')]"
)

//SCHEDULE-ACTION
int i = 1 //Boleh tukar based on nak click which one

TestObject scheduleActionIcon = new TestObject()
scheduleActionIcon.addProperty(
    "xpath",
    ConditionType.EQUALS,
    "(//img[contains(@src, 'ico-view.png') and @title='Contract Request'])[${i}]"
)

//AS&WHEN
TestObject actionIcon = new TestObject()
actionIcon.addProperty("xpath", ConditionType.EQUALS,
    "//img[contains(@src, 'ico-view.png') and @title='Contract Request']"
)

/* ==================================================
 * Click which one display
 * AS&WHEN || SCHEDULE
 * ================================================== */

boolean isSchedule = false //Detect not Schedule

/* =========================
 * CHECK & CLICK SCHEDULE
 * ========================= */
if (WebUI.waitForElementVisible(scheduleIcon, 10, FailureHandling.OPTIONAL)) {

    isSchedule = true //okay schedule -- proceed

    WebElement schedule = WebUiCommonHelper.findWebElement(scheduleIcon, 10)
    WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(schedule))

    if (WebUI.waitForElementVisible(scheduleActionIcon, 10, FailureHandling.OPTIONAL)) {

        WebElement scheduleAction = WebUiCommonHelper.findWebElement(scheduleActionIcon, 10)
        WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(scheduleAction))
    }
}

/* =================================
 * AS & WHEN (ONLY IF NOT SCHEDULE)
 * ================================= */
if (!isSchedule) { //If not schdule proceed AS&WHEN

    if (WebUI.waitForElementVisible(actionIcon, 10, FailureHandling.OPTIONAL)) {

        WebElement action = WebUiCommonHelper.findWebElement(actionIcon, 10)
        WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(action))
    }
}

/* =================================
 * GENERAL - Request Details
 * ================================= */
//Description
t(findTestObject('Object Repository/DP - Add To Cart/Contract List/Textarea Description'), Description)

//Tick to remain
c(findTestObject('Object Repository/DP - Add To Cart/Contract List/Tickbox - To remain'))

//Start Date
c(findTestObject('Object Repository/DP - Add To Cart/Contract List/Click Calender - Start Date'))
pickDate("2026-06-19")

//End Date
c(findTestObject('Object Repository/DP - Add To Cart/Contract List/Click Calender - End Date'))
pickDate("2026-07-18")

//Tick I declare
c(findTestObject('Object Repository/DP - Add To Cart/Contract List/Tickbox - I declare'))

//Supplier Branch Name
//selectDropdownByIndex(findTestObject('Object Repository/DP - Add To Cart/Contract List/Dropdown Supplier Branch'),ddSupplier)
def selectSuppBranchIfEnabled(String optionText) {
	
	TestObject suppBranchContainer = new TestObject('suppBranchContainer')
	suppBranchContainer.addProperty(
		'xpath',
		ConditionType.EQUALS,
		"//div[@id='flContractRequest_WAR_NGePportlet:form:suppBranch']"
		)
	
	TestObject suppBranchSelect = new TestObject('suppBranchHiddenSelect')
	suppBranchSelect.addProperty(
		'xpath',
		ConditionType.EQUALS,
		"//select[@id='flContractRequest_WAR_NGePportlet:form:suppBranch_input']"
		)
	
	WebUI.waitForElementPresent(suppBranchContainer, 10)
	
	String parentClass = WebUI.getAttribute(suppBranchContainer, 'class', FailureHandling.OPTIONAL)
	String disabledAttr = WebUI.getAttribute(suppBranchSelect, 'disabled', FailureHandling.OPTIONAL)
	
	boolean isDisabled =
		parentClass?.contains('ui-state-disabled') ||
		disabledAttr?.equalsIgnoreCase('disabled') ||
		disabledAttr?.equalsIgnoreCase('true')
	
	if (isDisabled) {
		println "Supplier Branch dropdown is disabled. Skipping selection."
		return
	}
	
	TestObject suppBranchTrigger = new TestObject('suppBranchTrigger')
	suppBranchTrigger.addProperty(
		'xpath',
		ConditionType.EQUALS,
		"//div[@id='flContractRequest_WAR_NGePportlet:form:suppBranch']/div[contains(@class,'ui-selectonemenu-trigger')]"
	)
	
	WebUI.waitForElementClickable(suppBranchTrigger, 10)
	WebUI.click(suppBranchTrigger)
	
	TestObject suppBranchOption = new TestObject('suppBranchOption_' + optionText)
	suppBranchOption.addProperty(
		'xpath',
		ConditionType.EQUALS,
		"//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[normalize-space(.)='${optionText}']"
	)
	
	WebUI.waitForElementVisible(suppBranchOption, 10)
	WebUI.click(suppBranchOption)
			println "Supplier Branch selected: " + optionText
}
