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
waitBlockUI(20)
// WebUI.delay(1)  // keep if you want; but overlay wait is usually enough
WebUI.delay(3)

t(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Username'), Username, 20)
waitBlockUI(20)
// WebUI.delay(1)  // keep if you want; but overlay wait is usually enough
WebUI.delay(0.5)

t(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Password'), Password, 20)
waitBlockUI(20)
// WebUI.delay(1)  // keep if you want; but overlay wait is usually enough
WebUI.delay(0.5)

c(findTestObject('Direct LOA/1. Direct LOA Requistioner/Login/Submit Username and Password'), 20)
waitBlockUI(30)
WebUI.delay(2)

WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

// TaskList
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Click Task List'), 20)
waitBlockUI(20)
WebUI.delay(1)

// Expand My Task
c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/span_My Task_ui-icon ui-icon-plusthick'), 20)
waitBlockUI(20)
WebUI.delay(1)

// Input Document Number
t(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Input Document Number'),
	Document_Number, 20
)
waitBlockUI(20)
WebUI.delay(1)

// Search
c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Search TaskList'), 20)
waitBlockUI(30)
WebUI.delay(2)

// Click TaskList Description
c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Click TaskList Description'), 20)
waitBlockUI(20)
/* =========================
 * Click Prepare LOA
 * ========================= */
WebUI.click(findTestObject('Object Repository/DLOA/8. Prepare LOA/Click Prepare LOA Button'))

t(findTestObject('Object Repository/DLOA/8. Prepare LOA/File Reference No_1'), FileReference1, 20)
t(findTestObject('Object Repository/DLOA/8. Prepare LOA/File Reference No_2'), FileReference2, 20)

c(findTestObject('Object Repository/DLOA/8. Prepare LOA/Date Picker COntract Details icon'), 20)

pickDate("2026-04-23")   // <-- put your date here
waitBlockUI(20)
WebUI.delay(1)

//Agreement Radio Button 

// value: 1 = Yes, 2 = No
def clickAgreementReq(def value) {

    WebUI.comment("DEBUG Aggrement raw value = [" + value + "]")
    WebUI.comment("DEBUG Aggrement class = " + (value == null ? "null" : value.getClass().getName()))

    if (value == null || value.toString().trim() == "") {
        assert false : "❌ Agreement value is empty"
    }

    String rawValue = value.toString().trim()
    int intValue = rawValue.toInteger()

    WebUI.comment("DEBUG Aggrement intValue = " + intValue)

    String label = (intValue == 1) ? "Yes" : "No"
    WebUI.comment("DEBUG Agreement label = " + label)

    TestObject opt = new TestObject("agreementReq_" + label)
    opt.addProperty("xpath", ConditionType.EQUALS,
        "//*[@id='_scDpLoaList_WAR_NGePportlet_:form:agreementReq']//label[normalize-space(.)='${label}']"
    )

    WebUI.waitForElementClickable(opt, 20)
    WebUI.click(opt)

    if (intValue == 2) {
        WebUI.comment("DEBUG value = 2, clicking popup Yes")
        c(findTestObject('Object Repository/DLOA/8. Prepare LOA/Click pop up Yes'), 20)
    }
}

clickAgreementReq(Aggrement)
/* =========================
 * LOA & Attachment
 * ========================= */
c(findTestObject('Object Repository/DLOA/8. Prepare LOA/LOA And Attachment/Side Menu LOA And Attachment'), 20)
waitBlockUI(30)
WebUI.delay(1)


c(findTestObject('Object Repository/DLOA/8. Prepare LOA/LOA And Attachment/Date Picker LOA and Attachment'), 20)
pickDate("2026-04-23")   // <-- put your date here
waitBlockUI(20)
WebUI.delay(1)

c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Button LOA Signer'), 20)
waitBlockUI(10)

// Username dropdown
selectDropdownByIndex(findTestObject('Object Repository/DLOA/8. Prepare LOA/LOA And Attachment/Dropdown Username'), 1)

t(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Input User Name'), LOASigner, 20)
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Search User Name'), 20)
waitBlockUI(30)

TestObject userRow = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LOA And Attachment Tab/Click in user Table')

wVisible(userRow, 20)
WebUI.scrollToElement(userRow, 2)
waitBlockUI(10)
WebUI.delay(1)

try {
	WebUI.doubleClick(userRow)
} catch (Exception e) {
	WebUI.delay(0.5)
	WebUI.click(userRow)
	WebUI.delay(0.15)
	WebUI.click(userRow)
}

c(findTestObject('Object Repository/DLOA/8. Prepare LOA/LOA And Attachment/Click upload Icon'), 3)
up(findTestObject('Object Repository/DLOA/8. Prepare LOA/LOA And Attachment/Choose File'),
	'C:\\Users\\nurul.atikah\\Documents\\File pdf_for testing.pdf',3)

c(findTestObject('Object Repository/DLOA/8. Prepare LOA/LOA And Attachment/Upload Icon LOA Signer Document'), 20)
waitBlockUI(30)

/* =========================
 * Save LOA (unchanged)
 * ========================= */
//WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Save LOA Application'))
c(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Submit and Save Button/Submit LOA Application'), 20)
waitBlockUI(30)

/* =========================
* WAIT LOADER + CAPTURE LOA MESSAGE (DYNAMIC LAxxxx) + APPEND TO EXCEL (SAME FILE)
* ========================= */

// ===== 1) Wait loader/blockUI gone (PrimeFaces common) =====
TestObject blockUI = new TestObject('blockUI')
blockUI.addProperty("xpath", ConditionType.EQUALS,
"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
)

if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
}

// ===== 2) Wait success message (global text; LOA number changes) =====
TestObject msgObj = new TestObject('msg_LOA_saved')
msgObj.addProperty("xpath", ConditionType.EQUALS,
"//span[contains(@class,'ui-messages-info-detail') and " +
"contains(.,'Letter of Acceptance (LOA)') and contains(.,'submitted successfully')]"
)

WebUI.waitForElementVisible(msgObj, 30)

// Wait until message text contains "LA"
String msg = ""
for (int i = 0; i < 15; i++) {
msg = WebUI.getText(msgObj, FailureHandling.OPTIONAL)
if (msg != null && msg.contains("LA")) break
WebUI.delay(1)
}

msg = (msg == null) ? "" : msg.trim()
WebUI.comment("Message: " + msg)

// ===== 3) Extract LA number dynamically =====
def matcher = (msg =~ /(LA\d+)/) // e.g. LA260000000000604
String loaNo = matcher.find() ? matcher.group(1) : ""

if (loaNo == "") {
WebUI.takeScreenshot()
assert false : "❌ LOA number not found. Message was: " + msg
}
WebUI.comment("✅ Captured LOA No: " + loaNo)

// ===== 4) Append to SAME Excel file (no timestamp file) =====
String filePath = "C:\\Users\\nurul.atikah\\Documents\\CDC - Work\\Automation\\Test Data\\Direct Purchased\\No SQ\\LOA_DP_CR_RN_LOA_PREPARE_LOA.xlsx"
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

// Close input stream to avoid Excel file lock
if (fis != null) fis.close()

// Next empty row
int nextRow = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : sheet.getLastRowNum() + 1
def row = sheet.createRow(nextRow)

row.createCell(0).setCellValue(now)
row.createCell(1).setCellValue(loaNo)
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
