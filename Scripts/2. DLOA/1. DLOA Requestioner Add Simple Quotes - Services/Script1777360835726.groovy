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

	/*===================================
	 function tick ikut index
	 ===================================*/
	def tickZoneTreeByIndex(int index) {
	  String xpath = "(//*[contains(@id,'treeZoneLocationPopup')]/span/div/div)[" + index + "]"
	  TestObject obj = new TestObject("zoneTreeTick_" + index)
	  obj.addProperty("xpath", ConditionType.EQUALS, xpath)
	
	  if (WebUI.verifyElementPresent(obj, 5, FailureHandling.OPTIONAL)) {
	
		  WebElement element = WebUiCommonHelper.findWebElement(obj, 10)
		  String checked = element.getAttribute("aria-checked")
	
		  // ONLY tick kalau belum tick
		  if (checked == null || checked != "true") {
	
			  WebUI.executeJavaScript("arguments[0].scrollIntoView(true);", Arrays.asList(element))
	
			  WebUI.executeJavaScript(
				  "arguments[0].click();",
				  Arrays.asList(element)
			  )
	
			  WebUI.delay(0.2)
		  }
	  }
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

/* =========================
 * DLOA - Requestioner
 * ========================= */
// Open Catalogue Search
	c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Catalogue Search'), 20)
	waitBlockUI(30)
	WebUI.delay(1)
		
	// Input Item Keyword
	TestObject itemKeyword = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Input Item Keyword')
	wVisible(itemKeyword, 20)
	WebUI.click(itemKeyword)
	WebUI.clearText(itemKeyword)
	WebUI.setText(itemKeyword, Keyword)
	WebUI.delay(0.5)
		
	// Input Supplier Name
	TestObject supplierName = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Input Supplier Name')
	wVisible(supplierName, 20)
	WebUI.click(supplierName)
	WebUI.clearText(supplierName)
	WebUI.setText(supplierName, SupplierName)
	WebUI.delay(0.5)
	
	// Input Item Code
	TestObject itemCode = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Input Item Code')
	wVisible(itemCode, 20)
	WebUI.click(itemCode)
	WebUI.clearText(itemCode)
	WebUI.setText(itemCode, ItemCode)
	WebUI.delay(0.5)
		
	// Click Search
	c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Button Search Supplier'), 20)
	waitBlockUI(30)
	WebUI.delay(1)
		
	// Click Action dropdown
	c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Dropdown Action - Simple Quote'), 20)
	waitBlockUI(20)
	WebUI.delay(0.5)
		
	// Click Add to Simple Quote
	c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Add to Simple Quote'), 20)
	waitBlockUI(30)
	WebUI.delay(1)

/* =========================
 * DLOA - General Information
 * ========================= */

TestObject title = findTestObject(
	'Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Title'
)
		
	String titleInput = DLAOTitle.toString().trim()
		
	WebUI.waitForElementVisible(title, 20)
	WebUI.waitForElementClickable(title, 20)
	WebUI.click(title)
	WebUI.delay(0.5)
		
	// clear like user
	WebUI.sendKeys(title, Keys.chord(Keys.CONTROL, 'a'))
	WebUI.delay(0.2)
	WebUI.sendKeys(title, Keys.chord(Keys.BACK_SPACE))
	WebUI.delay(0.3)
		
	// type slowly
	for (char ch : titleInput.toCharArray()) {
		WebUI.sendKeys(title, ch.toString())
		WebUI.delay(0.1)
	}


selectDropdownByIndex(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Dropdown Procurement Type Category'), ProcurementtypeCategory)

int rbType = Integer.parseInt(RBProcurementType.toString())

//Procurement Type Category
clickProcurementType(rbType)

// IF NOT 1 OR 3 → isi Reason + Justification
if (!(rbType == 1 || rbType == 3)) {

    selectDropdownByIndex(
        findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Dropdown Reason'),
        ReasonPK7
    )

    WebUI.setText(
        findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Justification'),
        Justification
    )
}

// Calendar
c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Start Date Picker icon'), 20)
WebUI.delay(1)

pickDate("2026-04-29")   // <-- put your date here
waitBlockUI(20)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Start Date Hour'),
	'20', true) // TIME (Hour)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Start Date Minute'),
	'15', true)  //TIME (Minutes)
WebUI.delay(1)

//c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/End Date Picker icon'), 20)
//WebUI.delay(1)

//pickDate("2026-03-4")   // <-- put your date here
waitBlockUI(20)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/End Date Hour'),
	'20', true) // TIME (Hour)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/End Date Minutes'),
	'15', true)
WebUI.delay(1) //TIME (Minutes)
/* =========================
 * Zone Item - Service
 * ========================= */
// Click Side Menu item
c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Side Menu item List'), 20)
waitBlockUI(30)
WebUI.delay(1)


/*=============================
 Looping For Service Item/Specification
===============================*/
// Set number of loops (1-10)
int loopCount = 2  // Set how many times you want the loop to run (1-10)

for (int i = 1; i <= loopCount; i++) {
    WebUI.comment("Loop #${i} of ${loopCount}")

    // First loop doesn't click "Action Item for Service"
    if (i > 1) {
        // Click Action Item for Service after the first loop starts
        c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Action Item for Service'), 20)
        waitBlockUI(20)
        WebUI.delay(1)
    }

    // Select the Fulfilment Type for Service - Ada condition, if select ONE OFF dgn Schedule, as and when
    selectDropdownByIndex(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Fulfilment Type Services'), Fulfilmenttype)
	
    // Input Specification for Service 1 (Dynamic Text + Loop Index)
	TestObject spec1Service = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Service Input Specification 1 TextBox')
	WebUI.waitForElementVisible(spec1Service, 20)
	String existingValue = WebUI.getAttribute(spec1Service, "value")
	
	if (existingValue == null || existingValue.trim().isEmpty()) {
	
		WebUI.waitForElementClickable(spec1Service, 20)
		WebUI.click(spec1Service)
		WebUI.clearText(spec1Service)
	
		String serviceSpec1WithLoopIndex = ServiceSpecification1 + i
		WebUI.setText(spec1Service, serviceSpec1WithLoopIndex)
	
		waitBlockUI(30)
		WebUI.delay(1)
	}

    // Input UOM for Service
    TestObject uomService = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Input UOM')
    wVisible(uomService, 20)
    WebUI.click(uomService)
    t(uomService, ServiceUOM, 20)
    WebUI.delay(1)
    WebUI.sendKeys(uomService, Keys.chord(Keys.ENTER))
    waitBlockUI(30)
    WebUI.delay(1)
	
	// FOR SCHEDULE & AS AND WHEN
	int type = Integer.parseInt(Fulfilmenttype.toString())

	if (type == 2 || type == 3) {
    TestObject freq = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Freq. Per UOM')
    TestObject duration = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Duration (Month)')

    // Tunggu element muncul dulu (optional supaya tak fail terus)
    if (WebUI.verifyElementPresent(freq, 10, FailureHandling.OPTIONAL)) {
        WebUI.waitForElementVisible(freq, 10)
        t(freq, Freq, 20)
        waitBlockUI(30)
        WebUI.delay(1)
    	}

    if (WebUI.verifyElementPresent(duration, 10, FailureHandling.OPTIONAL)) {
        WebUI.waitForElementVisible(duration, 10)
        t(duration, DurationMonth, 20)
        waitBlockUI(30)
        WebUI.delay(1)
		}
	}
	
    // Input Service Quantity
    t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Quantity'), ServiceQty, 20)
    waitBlockUI(30)
    WebUI.delay(1)

    // Input Service Period
    t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Service Period'), ServicePeriod, 20)
    waitBlockUI(30)
    WebUI.delay(1)

	// Click Action Item for Specification (Service)
	TestObject specBtn = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Action Item Click Specification')
	WebUI.waitForElementVisible(specBtn, 20)
	WebUI.waitForElementClickable(specBtn, 20)
	WebUI.click(specBtn)
	
	TestObject clickSpec = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Click Specification')
	
	WebUI.waitForElementClickable(clickSpec, 20)
	WebUI.click(clickSpec)
	waitBlockUI(20)
	WebUI.delay(1)

    // Input Specification 2 TextBox for Service
    TestObject spec2Service = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input Specification 2 TextBox')
    String serviceSpec2WithLoopIndex = ServiceSpecification2 + i  // Add loop index to Specification Text
    WebUI.setText(spec2Service, serviceSpec2WithLoopIndex) // type new value
    waitBlockUI(30)
    WebUI.delay(1)
}

// Convert dulu (PENTING untuk Test Suite)
int locationVal = Integer.parseInt(LocalityCoverage.toString())

// Click Side Menu
if (rbType == 1) {

    TestObject sideMenu = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/3. Criteria and Supplier List/Side Menu Criteria and Supplier List')
    WebUI.waitForElementClickable(sideMenu, 30)
    WebUI.click(sideMenu)

    waitBlockUI(30)
    WebUI.waitForPageLoad(30)
    WebUI.delay(1)

    /* ==================================================
     * Supplier Location
     * ================================================== */
    if (locationVal == 1) {

        TestObject locality = findTestObject(
            'Object Repository/DLOA/4. DLOA - Requestioner/3. Criteria and Supplier List/Locality Coverage'
        )

        WebUI.waitForElementVisible(locality, 20)
        WebUI.waitForElementClickable(locality, 20)
        selectDropdownByIndex(locality, locationVal)

        TestObject addLocationBtn = findTestObject(
            'Object Repository/DLOA/4. DLOA - Requestioner/3. Criteria and Supplier List/Click Button Add Location'
        )

        WebUI.click(addLocationBtn)
        waitBlockUI(10)

        def zoneGroups = [
            "Zone A": [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16]
        ]

        zoneGroups["Zone A"].each { idx ->
            tickZoneTreeByIndex(idx)
        }

        TestObject okBtn = findTestObject(
            'Object Repository/DLOA/4. DLOA - Requestioner/3. Criteria and Supplier List/Click Button OK'
        )

        WebUI.click(okBtn)
        waitBlockUI(10)
    }

    /* ==================================================
     * Supplier List
     * ================================================== */
    TestObject supplierListTab = findTestObject(
        'Object Repository/DLOA/4. DLOA - Requestioner/3. Criteria and Supplier List/Click Supplier List'
    )
    WebUI.click(supplierListTab)

    TestObject addSupplierBtn = findTestObject(
        'Object Repository/DLOA/4. DLOA - Requestioner/3. Criteria and Supplier List/Click Button Add Supplier'
    )

    TestObject supplierDropdown = findTestObject(
        'Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/DP - Supplier Dropdown'
    )

    TestObject supplierNameField = findTestObject(
        'Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Key In Business Name'
    )

    TestObject searchBtn = findTestObject(
        'Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Button Search Supplier'
    )

    TestObject supplierRow = findTestObject(
        'Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/DP - Click the Supplier'
    )

    TestObject selectSupplierBtn = findTestObject(
        'Object Repository/Direct LOA/1. Direct LOA Requistioner/Search Supplier From Requistioner/Select Supplier'
    )

    // 🔥 LOOP ADD SUPPLIER 2 KALI
    for (int i = 1; i <= 2; i++) {

        WebUI.waitForElementClickable(addSupplierBtn, 30)
        WebUI.click(addSupplierBtn)

        waitBlockUI(30)
        WebUI.waitForPageLoad(30)

        selectDropdownByIndex(supplierDropdown, 1)

        WebUI.setText(supplierNameField, AddSupplierNameNormal)
        WebUI.click(searchBtn)

        waitBlockUI(30)

        WebUI.waitForElementVisible(supplierRow, 20)
        WebUI.scrollToElement(supplierRow, 2)

        try {
            WebUI.click(supplierRow)
        } catch (Exception e) {
            WebUI.enhancedClick(supplierRow, FailureHandling.OPTIONAL)
        }

        WebUI.doubleClick(supplierRow, FailureHandling.OPTIONAL)

        WebUI.click(selectSupplierBtn)

        waitBlockUI(10)
        WebUI.delay(2)
    }

} else {
    // optional logic kalau bukan type 1
}

/* =========================
 * Save LOA (unchanged)
 * ========================= */
//WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Save LOA Application'))
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Submit LOA Application'))
WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Confirmation Pop up After Submit'))

waitBlockUI(30)

/* =========================
 * WAIT LOADER + CAPTURE SQ MESSAGE (DYNAMIC SQxxxx) + APPEND TO EXCEL (SAME FILE)
 * ========================= */

// ===== 1) Wait loader/blockUI gone (PrimeFaces common) =====
TestObject blockUI = new TestObject('blockUI')
blockUI.addProperty("xpath", ConditionType.EQUALS,
	"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
)

if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
	WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
}

// ===== 2) Wait success message (global text; SQ number changes) =====
TestObject msgObj = new TestObject('msg_SQ_saved')
msgObj.addProperty("xpath", ConditionType.EQUALS,
	"//span[contains(@class,'ui-messages-info-detail') and " +
	"contains(.,'Simple Quote') and contains(.,'is successfully submitted')]"
)

WebUI.waitForElementVisible(msgObj, 30)

// Wait until message text contains "SQ"
String msg = ""
for (int i = 0; i < 15; i++) {
	msg = WebUI.getText(msgObj, FailureHandling.OPTIONAL)
	if (msg != null && msg.contains("SQ")) break
	WebUI.delay(1)
}

msg = (msg == null) ? "" : msg.trim()
WebUI.comment("Message: " + msg)

// ===== 3) Extract SQ number dynamically =====
def matcher = (msg =~ /(SQ\d+)/)   // e.g. SQ260000000000604
String sqNo = matcher.find() ? matcher.group(1) : ""

if (sqNo == "") {
	WebUI.takeScreenshot()
	assert false : "❌ SQ number not found. Message was: " + msg
}
WebUI.comment("✅ Captured SQ No: " + sqNo)

// ===== 4) Append to SAME Excel file (no timestamp file) =====
String filePath = "C:\\Users\\nurul.atikah\\Documents\\CDC - Work\\Automation\\Test Data\\Direct Purchased\\No SQ\\FL_DP_CR_DLOA_Requestioner_Service_NEW.xlsx"
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
	header.createCell(1).setCellValue("SQ No")
	header.createCell(2).setCellValue("Message")
}

// Close input stream to avoid Excel file lock
if (fis != null) fis.close()

// Next empty row
int nextRow = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : sheet.getLastRowNum() + 1
def row = sheet.createRow(nextRow)

row.createCell(0).setCellValue(now)
row.createCell(1).setCellValue(sqNo)
row.createCell(2).setCellValue(msg)

// Save back to SAME file
FileOutputStream fos = new FileOutputStream(filePath)
wb.write(fos)
fos.close()
wb.close()

WebUI.comment("✅ Appended to Excel: " + filePath)

/* =========================
 * Sign Out
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))

// wait until logout is completed (choose one)
WebUI.waitForPageLoad(20)
WebUI.closeBrowser()
