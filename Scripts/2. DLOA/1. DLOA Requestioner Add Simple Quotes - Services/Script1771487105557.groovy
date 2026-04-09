import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.webui.common.WebUiCommonHelper
import com.kms.katalon.core.webui.driver.DriverFactory
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI

import org.openqa.selenium.Keys
import org.openqa.selenium.JavascriptExecutor
import org.openqa.selenium.WebDriver
import org.openqa.selenium.chrome.ChromeDriver
import org.openqa.selenium.chrome.ChromeOptions

import java.nio.file.Files
import java.nio.file.Paths
import java.text.SimpleDateFormat

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream

import com.kms.katalon.core.testobject.ConditionType

/* =========================
 * Calendar Picker Date
 * ========================= */

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
/* =========================
 * Helpers (KEEP YOUR LOGIC, ONLY ADD WAITS)
 * ========================= */

// Convert excel/csv value to int safely: "0", 0, 0.0, "1.0"
int toInt(def v, int defaultVal = 0) {
	if (v == null) return defaultVal
	return new BigDecimal(v.toString().trim()).intValue()
}

// PrimeFaces overlay wait (your original)
def waitBlockUI(int timeout = 30) {
	TestObject blockUI = new TestObject('blockUI')
	blockUI.addProperty("xpath", ConditionType.EQUALS,
		"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
	)

	if (WebUI.verifyElementPresent(blockUI, 1, FailureHandling.OPTIONAL)) {
		WebUI.waitForElementNotVisible(blockUI, timeout, FailureHandling.OPTIONAL)
	}
}

/* ---------- NEW: Lightweight wait wrappers (non-invasive) ---------- */
def wVisible(TestObject obj, int timeout = 1) {
	waitBlockUI(Math.min(timeout, 1))
	WebUI.waitForElementVisible(obj, timeout, FailureHandling.STOP_ON_FAILURE)
}

def wClickable(TestObject obj, int timeout = 1) {
	wVisible(obj, timeout)
	WebUI.waitForElementClickable(obj, timeout, FailureHandling.STOP_ON_FAILURE)
}

def c(TestObject obj, int timeout = 1) { // click with wait + tiny retry
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
	// last try (keep WebUI.click)
	wClickable(obj, timeout)
	WebUI.click(obj)
	waitBlockUI(1)
}

def dc(TestObject obj, int timeout = 1) { // double click with wait
	try {
		wClickable(obj, timeout)
		WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
		WebUI.doubleClick(obj, FailureHandling.OPTIONAL)
		waitBlockUI(1)
	} catch (Exception e) {
		// keep optional behavior
		WebUI.doubleClick(obj, FailureHandling.OPTIONAL)
		waitBlockUI(1)
	}
}

def t(TestObject obj, def value, int timeout = 1) { // setText with wait
	wVisible(obj, timeout)
	WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
	WebUI.setText(obj, (value == null ? "" : value.toString()))
}

def up(TestObject obj, String filePath, int timeout = 1) { // upload with wait
	wVisible(obj, timeout)
	WebUI.uploadFile(obj, filePath)
	waitBlockUI(1)
}

/* =========================
 * open PrimeFaces dropdown (your original + add waits)
 * ========================= */
def openPFDropdown(TestObject triggerObj) {

	TestObject panelOpen = new TestObject('pfPanelOpen')
	panelOpen.addProperty("xpath", ConditionType.EQUALS,
		"//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]"
	)

	// was: WebUI.click(triggerObj)
	c(triggerObj, 20)
	WebUI.delay(0.3)

	if (!WebUI.waitForElementVisible(panelOpen, 1, FailureHandling.OPTIONAL)) {
		c(triggerObj, 20)
		WebUI.delay(0.3)
		WebUI.waitForElementVisible(panelOpen, 3, FailureHandling.OPTIONAL)
	}
}

// click PrimeFaces option by index (0-based) (your original + use click wrapper)
def clickPFOptionByIndex(int index0) {
	TestObject opt = new TestObject("pfOpt_" + index0)
	opt.addProperty("xpath", ConditionType.EQUALS,
		"(//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[contains(@class,'ui-selectonemenu-item')])[${index0 + 1}]"
	)

	// was: waitForElementClickable + click
	c(opt, 20)
	WebUI.delay(0.2)
	waitBlockUI(20)
}

/**
 * Universal dropdown select by index (SAFE): (your original)
 */
def selectDropdownByIndex(TestObject dropdownObj, def indexFromData) {

	int idx0 = toInt(indexFromData) // NO -1 because data already 0-based

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
			WebUI.delay(0.5) // PrimeFaces re-render
		}
	}

	assert false : "❌ Dropdown failed (stale/DOM refresh): " + dropdownObj.getObjectId()
}

/* =========================
 * Browser Setup (Guest + Clean) (UNCHANGED)
 * ========================= */
String userDataDir = Files.createTempDirectory('katalon-clean').toString()

ChromeOptions options = new ChromeOptions()
options.addArguments('--guest')
options.addArguments('--incognito')
options.addArguments('--user-data-dir=' + userDataDir)
options.addArguments('--disable-features=PasswordLeakDetection,PasswordManagerOnboarding')
options.addArguments('--disable-save-password-bubble')
options.addArguments('--no-first-run')
options.addArguments('--no-default-browser-check')
options.setExperimentalOption('prefs', [
	('credentials_enable_service') : false,
	('profile.password_manager_enabled') : false,
	('profile.default_content_setting_values.notifications') : 2
])

WebDriver driver = new ChromeDriver(options)
DriverFactory.changeWebDriver(driver)

/* =========================
 * Test Flow (SAME FLOW, ONLY WRAP ACTIONS)
 * ========================= */
WebUI.navigateToUrl('http://ngepsit.eperolehan.com.my/home')
WebUI.maximizeWindow()
waitBlockUI(20)

// Language
wVisible(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 20)
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)
waitBlockUI(20)
// WebUI.delay(1)  // keep if you want; but overlay wait is usually enough
WebUI.delay(1)

// Login
c(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Right Top Menu Login'), 20)
t(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Username'), Username, 20)
t(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Password'), Password, 20)
c(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Submit Username and Password'), 20)
waitBlockUI(30)
WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

/* =========================
 * DLOA - Requestioner
 * ========================= */

WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Catalogue Search'))

//WebUI.setText(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Input Item Keyword'),Keyword)

WebUI.setText(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Input Supplier Name'),
	SupplierName)

WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Button Search Supplier'))

WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Dropdown Action - Simple Quote'))

WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Add to Simple Quote'))

/* =========================
 * DLOA - General Information
 * ========================= */

WebUI.setText(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Title'),
	DLAOTitle)

selectDropdownByIndex(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Dropdown Procurement Type Category'), Procurementtype)

WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Procurement Type Radio Button'))

selectDropdownByIndex(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Dropdown Reason'), ReasonPK7)

WebUI.setText(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Justification'),
	Justification)

// Calendar
c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Start Date Picker icon'), 20)
WebUI.delay(1)

pickDate("2026-03-17")   // <-- put your date here
waitBlockUI(20)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Start Date Hour'),
	'10', true)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Start Date Minute'),
	'15', true)
WebUI.delay(1)

//pickDate("2026-03-4")   // <-- put your date here
waitBlockUI(20)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/End Date Hour'),
	'10', true)
WebUI.delay(1)

WebUI.selectOptionByValue(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/End Date Minutes'),
	'30', true)
WebUI.delay(1)
/* =========================
 * Zone Item - Service
 * ========================= 
c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Side Menu item List'), 20)
waitBlockUI(30)
WebUI.delay(1)

/*TestObject AddService = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Action Item for Service')
WebUI.scrollToElement(AddService, 5)
wClickable(AddService, 20)
WebUI.click(AddService)
waitBlockUI(20)

t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Service Input Specification 1 TextBox'),
	ServiceSpecification1, 20
)*/

/*selectDropdownByIndex(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Fulfilment Type Services'), Fulfilmenttype)


TestObject uomService = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Input UOM')
wVisible(uomService, 20)
WebUI.click(uomService)
t(uomService, ServiceUOM, 20)
WebUI.delay(1)
WebUI.sendKeys(uomService, Keys.chord(Keys.ENTER))
WebUI.delay(1)

t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Freq. per UOM'),
	Freq, 20
)

TestObject durationObj = findTestObject(
  'Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Duration (Month)'
)

waitBlockUI(20)
WebUI.waitForElementClickable(durationObj, 20)
WebUI.click(durationObj)

// clear then input
WebUI.sendKeys(durationObj, Keys.chord(Keys.CONTROL, 'a'))
WebUI.sendKeys(durationObj, Keys.chord(Keys.BACK_SPACE))

WebUI.sendKeys(durationObj, DurationMonth)   // e.g. "60"
waitBlockUI(20)

t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Quantity'),
	ServiceQty, 20
)
WebUI.delay(1)

t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Service Period'),
	ServicePeriod, 20
	
)
WebUI.delay(1)

c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Action Item Click Specification'), 20)
def spec = findTestObject('Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Click Specification')
wVisible(spec, 20)
def el = WebUiCommonHelper.findWebElement(spec, 10)
((JavascriptExecutor) DriverFactory.getWebDriver()).executeScript("arguments[0].click();", el)

t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Input Specification 2 TextBox'),
	ServiceSpecification2, 20
)*/

/* =========================
 * Zone Item - Service 2 Item and Above
 * ========================= */
// Click Side Menu item
c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Side Menu item List'), 20)
waitBlockUI(30)
WebUI.delay(1)

// Set number of loops (1-10)
int loopCount = 1  // Set how many times you want the loop to run (1-10)

for (int i = 1; i <= loopCount; i++) {
	WebUI.comment("Loop #${i} of ${loopCount}")

	// First loop doesn't click "Action Item for Service"
	if (i > 1) {
		// Click Action Item for Service after the first loop starts
		c(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/1. General Information/Click Action Item for Service'), 20)
		waitBlockUI(20)
		WebUI.delay(1)
	}
	// Input Specification for Service 1 (Dynamic Text + Loop Index)
	TestObject spec1Service = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Service Input Specification 1 TextBox')
	WebUI.waitForElementVisible(spec1Service, 20)
	WebUI.waitForElementClickable(spec1Service, 20)
	WebUI.click(spec1Service)                 // focus
	WebUI.clearText(spec1Service)             // clear existing value
	String serviceSpec1WithLoopIndex = ServiceSpecification1 + i  // Add loop index to Specification Text
	WebUI.setText(spec1Service, serviceSpec1WithLoopIndex) // type new value
	waitBlockUI(30)
	WebUI.delay(2)
	
	// Select the Fulfilment Type for Service
	selectDropdownByIndex(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Fulfilment Type Services'), Fulfilmenttype)
	waitBlockUI(30)
	WebUI.delay(2)
	

	// Input UOM for Service
	TestObject uomService = findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Input UOM')
	wVisible(uomService, 20)
	WebUI.click(uomService)
	t(uomService, ServiceUOM, 20)
	WebUI.delay(2)
	WebUI.sendKeys(uomService, Keys.chord(Keys.ENTER))
	waitBlockUI(30)
	WebUI.delay(1)
	
	t(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Freq. per UOM'),Freq, 20)
	waitBlockUI(30)
	WebUI.delay(1)
	
	TestObject durationObj = findTestObject(
	  'Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Duration (Month)'
	)
	waitBlockUI(20)
	WebUI.waitForElementClickable(durationObj, 20)
	WebUI.click(durationObj)
	
	// clear then input
	WebUI.sendKeys(durationObj, Keys.chord(Keys.CONTROL, 'a'))
	WebUI.sendKeys(durationObj, Keys.chord(Keys.BACK_SPACE))
	
	WebUI.sendKeys(durationObj, DurationMonth)   // e.g. "60"
	waitBlockUI(30)
	WebUI.delay(1)


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
	WebUI.delay(2)

	// Input Specification 2 TextBox for Service
	TestObject spec2Service = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input Specification 2 TextBox')
	String serviceSpec2WithLoopIndex = ServiceSpecification2 + i  // Add loop index to Specification Text
	WebUI.setText(spec2Service, serviceSpec2WithLoopIndex) // type new value
	waitBlockUI(30)
	WebUI.delay(3)
}

/* =========================
 * Save LOA (unchanged)
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Save LOA Application'))
//WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Submit LOA Application'))
//WebUI.click(findTestObject('Object Repository/DLOA/4. DLOA - Requestioner/2. Item List/Confirmation Pop up After Submit'))

waitBlockUI(30)


/* =========================
 * WAIT LOADER + CAPTURE SQ MESSAGE (DYNAMIC SQxxxx) + APPEND TO EXCEL (SAME FILE)
 * =========================  */

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
	"contains(.,'Simple Quote') and contains(.,'is successfully submitted.')]"
)

WebUI.waitForElementVisible(msgObj, 20)

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
def matcher = (msg =~ /(SQ\d+)/)
String sqNo = matcher.find() ? matcher.group(1) : ""

if (sqNo == "") {
	WebUI.takeScreenshot()
	assert false : "❌ SQ number not found. Message was: " + msg
}
WebUI.comment("✅ Captured SQ No: " + sqNo)

// ===== 4) Append to SAME Excel file (no timestamp file) =====
String filePath = "C:\\Users\\hadishafiq\\Desktop\\PrepData\\fAHMI DP.xlsx"
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
 * =========================*/
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))

// wait until logout is completed (choose one)
WebUI.waitForPageLoad(20)
WebUI.closeBrowser()
