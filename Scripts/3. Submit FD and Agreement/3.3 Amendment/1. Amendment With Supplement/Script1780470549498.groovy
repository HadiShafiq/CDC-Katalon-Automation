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
 * 7) CALENDAR PICKER DATE -- khas untuk AMENDMENT SAHAJA
 * ========================================================= */
	
def pickDate(String yyyyMmDd) {

    // =========================
    // a. OPEN DATEPICKER FIRST
    // =========================
    TestObject openCalendarBtn = new TestObject('openCalendarBtn')
    openCalendarBtn.addProperty("xpath", ConditionType.EQUALS,
        "//button[contains(@class,'ui-datepicker-trigger')]"
    )

    WebUI.waitForElementClickable(openCalendarBtn, 10)
    WebUI.click(openCalendarBtn)

    // =========================
    // b. WAIT FOR DATEPICKER
    // =========================
    TestObject dp = new TestObject('dp')
    dp.addProperty("xpath", ConditionType.EQUALS,
        "//*[@id='ui-datepicker-div' and not(contains(@style,'display: none'))]"
    )
    WebUI.waitForElementVisible(dp, 20)

    // =========================
    // c. PARSE DATE
    // =========================
    def parts = yyyyMmDd.split('-')
    int targetYear = parts[0] as int
    int targetMonthIndex = (parts[1] as int) - 1
    String targetDay = String.valueOf(parts[2] as int)

    // =========================
    // d. NAV BUTTONS
    // =========================
    TestObject nextBtn = new TestObject('dpNext')
    nextBtn.addProperty("xpath", ConditionType.EQUALS,
        "//*[@id='ui-datepicker-div']//a[contains(@class,'ui-datepicker-next')]"
    )

    TestObject prevBtn = new TestObject('dpPrev')
    prevBtn.addProperty("xpath", ConditionType.EQUALS,
        "//*[@id='ui-datepicker-div']//a[contains(@class,'ui-datepicker-prev')]"
    )

    // =========================
    // e. LOOP MONTHS
    // =========================
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

        // kalau target belum ada → next month
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
 * ========================= */
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

/* ==========================
 * CLICK CONTRACT MAINTENANCE
 * ==========================*/
c(findTestObject('Object Repository/FD and Agreement/CM - Amendment/Click Contract Maintenance'))

t(findTestObject('Object Repository/FD and Agreement/CM - Amendment/Contract No'), ContractNo)

c(findTestObject('Object Repository/DP - Add To Cart/Payment Match/Click Button Search'))

c(findTestObject('Object Repository/DP - Add To Cart/Payment Match/Click Option'))

c(findTestObject('Object Repository/DP - Add To Cart/Payment Match/Amendment With Supplement'))

/* ==========================
 * AMENDMENT DETAIL 
 * ==========================*/
c(findTestObject('Object Repository/DP - Add To Cart/Payment Match/Menu Amendment Detail'))

/* ====================================================
 * TICK BOX FOR - Amendment Type & Approving Authority
 * 1-7 Amendment Type 
 * 8-9 Approving Authority
 * ====================================================*/
def tickBox(indexValue, dropdownValue = null) {

    def selectedIndexes = indexValue.toString().split(',').collect { it.trim() }

    selectedIndexes.each { idx ->

        int indexNo = idx.toInteger()

        TestObject checkbox = new TestObject("checkbox_" + indexNo)
        checkbox.addProperty("xpath", ConditionType.EQUALS,
            "(//div[contains(@class,'ui-chkbox-box')])[" + indexNo + "]"
        )

        WebUI.waitForElementVisible(checkbox, 10)
        WebUI.scrollToElement(checkbox, 5)

        String classAttr = WebUI.getAttribute(checkbox, "class")

        if (classAttr == null || !classAttr.contains("ui-state-active")) {
            WebUI.click(checkbox)
        }

        // =============================================================
        // INDEX 9 SPECIAL FLOW - Will display Desk Officer Committee
        // =============================================================
        if (indexNo == 9 && dropdownValue != null) {

            TestObject dropdownTrigger = new TestObject("dropdown_trigger")
            dropdownTrigger.addProperty("xpath", ConditionType.EQUALS,
                "//div[contains(@class,'ui-selectonemenu-trigger')]"
            )

            WebUI.waitForElementClickable(dropdownTrigger, 10)
            WebUI.click(dropdownTrigger)

            TestObject option = new TestObject("dropdown_option_" + dropdownValue)
            option.addProperty("xpath", ConditionType.EQUALS,
                "//li[contains(normalize-space(),'" + dropdownValue + "')]"
            )

            WebUI.waitForElementVisible(option, 10)
            WebUI.click(option)
        }
    }
}
	
//Call Function: Tick , Desk Officer Committee
tickBox(tickAmendmentnAuthority, deskOfficer)

/* =========================
 * FOR FILE REFRENCES NO.
 * ------------------------
 * Get LOANo. and  CTNo.
 * =========================*/
TestObject loaNo = new TestObject('loaNo')

loaNo.addProperty(
	"xpath",
	ConditionType.EQUALS,
	"//label[normalize-space()='Letter of Acceptance (LOA) No.']/ancestor::td/following-sibling::td[1]//a"
)

WebUI.waitForElementVisible(loaNo, 20)

String loaNum = WebUI.getText(loaNo).trim()

println("LOA No = " + loaNum)

/* ============================
 * Prepare Purchase Order No
 * ============================*/
String loaNumber = "FILE REF - " + loaNum

/* =========================
 * File References Field
 * =========================*/
TestObject fileRef = new TestObject('fileRef')

fileRef.addProperty(
	"xpath",
	ConditionType.EQUALS,
	"(//input[contains(@id,'form') and @type='text'])[1]"
)
/* ====================================
 * Key In Input Payment Description
 * ====================================*/
t(fileRef, loaNumber)
WebUI.delay(2)

/* =======================================================
 * XPATH FOR ALL INPUT 
 * =======================================================*/
//Months
TestObject monthsField = new TestObject()
monthsField.addProperty("xpath", ConditionType.EQUALS,
	"//input[contains(@id,'extendMonthsId')]"
)

//Day
TestObject daysField = new TestObject()
daysField.addProperty("xpath", ConditionType.EQUALS,
	"//input[contains(@id,'extendDaysId')]"
)

//Additional Contract Amount (RM)
TestObject addContAmtField = new TestObject()
addContAmtField.addProperty("xpath", ConditionType.EQUALS,
	"//input[contains(@id,'addValueTxt')]"
)

//Effective Date
TestObject openCalendarBtn = new TestObject('openCalendarBtn')
openCalendarBtn.addProperty("xpath", ConditionType.EQUALS,
	"//button[contains(@class,'ui-datepicker-trigger')]"
)

//New GST Amount (RM)
TestObject newGSTField = new TestObject()
newGSTField.addProperty("xpath", ConditionType.EQUALS,
	"//input[contains(@id,'newGstAmtTxt')]"
)

//New Sales Tax/Service Tax Amount (RM)	
TestObject newTaxField = new TestObject()
newTaxField.addProperty("xpath", ConditionType.EQUALS,
	"//input[contains(@id,'newStaxAmtTxt')]"
)

/* ==================================================
 * Fill Display A Fields [1]
 * (Months & Days)
 * ==================================================*/
//Months 
if (WebUI.waitForElementVisible(monthsField, 10)) {

    WebElement monthEl = WebUiCommonHelper.findWebElement(monthsField, 10)

    WebUI.executeJavaScript("arguments[0].focus();", Arrays.asList(monthEl))
    WebUI.executeJavaScript("arguments[0].value='';", Arrays.asList(monthEl))
    WebUI.executeJavaScript("arguments[0].value = arguments[1];", Arrays.asList(monthEl, monthValue)) //baca dr variable

    WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('input'));", Arrays.asList(monthEl))
    WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('change'));", Arrays.asList(monthEl))
    WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('blur'));", Arrays.asList(monthEl))
}

//Days
if (WebUI.waitForElementVisible(daysField, 10)) {
	
	WebElement dayEl = WebUiCommonHelper.findWebElement(daysField, 10)
	
	WebUI.executeJavaScript("arguments[0].focus();", Arrays.asList(dayEl))
	WebUI.executeJavaScript("arguments[0].value='';", Arrays.asList(dayEl))
	WebUI.executeJavaScript("arguments[0].value = arguments[1];", Arrays.asList(dayEl, dayValue))
	
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('input'));", Arrays.asList(dayEl))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('change'));", Arrays.asList(dayEl))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('blur'));", Arrays.asList(dayEl))
}

/* ========================================================================================
 * Fill Display B Fields [2,3,4]
 * (Additional Contract Amount (RM) , Effective Date)
 * ----------------------------------------------------------------------------------------
 * Fill Display C Fields [6]
 * (Additional Contract Amount (RM) ,Effective Date, GST, New Tax)
 * ----------------------------------------------------------------------------------------
 * Fill Display C Fields [5,7]
 * (Effective Date)
 * ========================================================================================*/
//Additional Contract Amount (RM)
if (WebUI.waitForElementVisible(addContAmtField, 10)) {
	
	WebElement addEl = WebUiCommonHelper.findWebElement(addContAmtField, 10)
	
	WebUI.executeJavaScript("arguments[0].focus();", Arrays.asList(addEl))
	WebUI.executeJavaScript("arguments[0].value='';", Arrays.asList(addEl))
	WebUI.executeJavaScript("arguments[0].value = arguments[1];", Arrays.asList(addEl, amountValue))
	
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('input'));", Arrays.asList(addEl))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('change'));", Arrays.asList(addEl))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('blur'));", Arrays.asList(addEl))
}

//Effective Date - Calender
if (WebUI.waitForElementVisible(openCalendarBtn, 10)) {
	
	WebElement btnEl = WebUiCommonHelper.findWebElement(openCalendarBtn, 10)
	
	WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(btnEl))
	
	pickDate(dateValue)
}

//GST
if (WebUI.waitForElementVisible(newGSTField, 10)) {
		
	WebElement gstF = WebUiCommonHelper.findWebElement(newGSTField, 10)
		
	WebUI.executeJavaScript("arguments[0].focus();", Arrays.asList(gstF))
	WebUI.executeJavaScript("arguments[0].value='';", Arrays.asList(gstF))
	WebUI.executeJavaScript("arguments[0].value = arguments[1];", Arrays.asList(gstF, gstValue))
		
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('input'));", Arrays.asList(gstF))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('change'));", Arrays.asList(gstF))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('blur'));", Arrays.asList(gstF))
}

//New Tax
if (WebUI.waitForElementVisible(newTaxField, 10)) {
		
	WebElement taxF = WebUiCommonHelper.findWebElement(newTaxField, 10)
		
	WebUI.executeJavaScript("arguments[0].focus();", Arrays.asList(taxF))
	WebUI.executeJavaScript("arguments[0].value='';", Arrays.asList(taxF))
	WebUI.executeJavaScript("arguments[0].value = arguments[1];", Arrays.asList(taxF, taxValue))
		
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('input'));", Arrays.asList(taxF))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('change'));", Arrays.asList(taxF))
	WebUI.executeJavaScript("arguments[0].dispatchEvent(new Event('blur'));", Arrays.asList(taxF))
}

/* ==================================================
 * Approval Letter - For Upload File
 * ==================================================*/
TestObject uploadIcon = findTestObject('Object Repository/FD and Agreement/CM - Amendment/Upload Button')
TestObject uploadBtn  = findTestObject('Object Repository/FD and Agreement/CM - Amendment/Upload File Button')

if (WebUI.waitForElementVisible(uploadIcon, 5, FailureHandling.OPTIONAL)) {

	WebUI.click(uploadIcon)

	//Wait for popup Upload File
	WebUI.delay(2)

	// create fresh object everytime upload
	TestObject fileInput = new TestObject()
	fileInput.addProperty(
		"xpath",
		ConditionType.EQUALS,
		"//input[contains(@id,'approvalLetterId_input')]"
	)

	WebUI.waitForElementPresent(fileInput, 10)

	WebUI.uploadFile(fileInput, uploadFilePath)

	if (WebUI.waitForElementClickable(uploadBtn, 20, FailureHandling.OPTIONAL)) {
		WebUI.click(uploadBtn)
	}

	waitBlockUI(30)

} else {
	WebUI.comment("Upload section NOT visible → skip upload")
}
