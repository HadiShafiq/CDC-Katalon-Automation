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
waitBlockUI(20)
WebUI.delay(0.5)


c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/MyTask_Tasklist_Dropdown'))

selectDropdownByIndex(findTestObject('Object Repository/FD and Agreement/Agreement Application/Common TaskList Funtion/MyTask DocumentType Dropdown'), DocumentType)
waitBlockUI(20)
WebUI.delay(0.5)

//Input Document Number
t(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Input Document Number'),Document_Number)
waitBlockUI(20)
WebUI.delay(0.5)


c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Search TaskList'))

// wait loader gone
waitBlockUI(30)

// wait table/data loaded (VERY IMPORTANT)
TestObject table = new TestObject('taskTable')
table.addProperty("xpath", ConditionType.EQUALS,
	"//tbody[contains(@id,'taskListGroupId_data')]"
)

WebUI.waitForElementVisible(table, 20)

// small buffer
WebUI.delay(1)

//Click TaskList Description
c(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Click TaskList Description'))
waitBlockUI(20)
WebUI.delay(0.5)

t(findTestObject('Object Repository/FD and Agreement/FD Application/Fulfilment Details/Physical Contract No'),PhysicalContractNo)
waitBlockUI(20)
WebUI.delay(0.5)

t(findTestObject('Object Repository/FD and Agreement/FD Application/Fulfilment Details/Service Period'),ServicePeriod)
waitBlockUI(20)
WebUI.delay(0.5)

c(findTestObject('Object Repository/FD and Agreement/Side Menu/FD Application/Agency Side Menu'))
waitBlockUI(20)
WebUI.delay(0.5)


int agencyLoopCount = 3

for (int i = 1; i <= agencyLoopCount; i++) {

	// =========================
	// Get data by loop
	// =========================
	def ministryValue = binding.getVariable("Ministry${i}")
	def jabatanValue  = binding.getVariable("Jabatan${i}")
	def ptjCodeValue  = binding.getVariable("PTJCode${i}")

	// =========================
	// Add Agency
	// =========================
	c(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Add Button Agency'))
	waitBlockUI(20)
	WebUI.delay(0.5)

	// =========================
	// Select Ministry
	// =========================
	selectDropdownByIndex(
		findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Ministry Dropdown'),
		ministryValue
	)
	waitBlockUI(20)
	WebUI.delay(0.5)

	// =========================
	// Select Jabatan
	// =========================
	selectDropdownByIndex(
		findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Jabatan Dropdown'),
		jabatanValue
	)
	waitBlockUI(20)
	WebUI.delay(2)

	// =========================
	// Input PTJ Code
	// =========================
	TestObject safeArea = new TestObject("safeArea_Header_${i}")
	safeArea.addProperty(
		"xpath",
		ConditionType.EQUALS,
		"//*[@id='_ctFulfilmentDetail_WAR_NGePportlet_:form:agencyPopUpId_header']/span"
	)

	TestObject ptjCode = findTestObject(
		'Object Repository/FD and Agreement/FD Application/Agency/PTJ Code'
	)

	String ptjCodeInput = ptjCodeValue.toString().trim()

	WebUI.waitForElementVisible(ptjCode, 20)
	WebUI.waitForElementClickable(ptjCode, 20)
	WebUI.click(ptjCode)
	WebUI.delay(0.5)

	WebUI.sendKeys(ptjCode, Keys.chord(Keys.CONTROL, 'a'))
	WebUI.delay(0.3)
	WebUI.sendKeys(ptjCode, Keys.chord(Keys.BACK_SPACE))
	WebUI.delay(0.5)

	for (char ch : ptjCodeInput.toCharArray()) {
		WebUI.sendKeys(ptjCode, ch.toString())
		WebUI.delay(0.15)
	}

	WebUI.click(safeArea, FailureHandling.OPTIONAL)

	waitBlockUI(10)
	WebUI.delay(1)

	String finalPTJCode = WebUI.getAttribute(ptjCode, 'value')
	println("Loop ${i} Final PTJ Code = " + finalPTJCode)

	// =========================
	// Search Agency
	// =========================
	c(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Search Agency'))
	waitBlockUI(20)
	WebUI.delay(0.5)

	// =========================
	// Select First Agency Row
	// =========================
	TestObject firstAgencyRow = new TestObject("firstAgencyRow_${i}")
	firstAgencyRow.addProperty(
		"xpath",
		ConditionType.EQUALS,
		"//*[@id='_ctFulfilmentDetail_WAR_NGePportlet_:form:tableAgency_data']/tr[1]"
	)

	if (!WebUI.waitForElementVisible(firstAgencyRow, 10, FailureHandling.OPTIONAL)) {
		WebUI.delay(2)
	}

	WebUI.waitForElementClickable(firstAgencyRow, 20)
	WebUI.scrollToElement(firstAgencyRow, 2)

	try {
		WebUI.click(firstAgencyRow)
	} catch (Exception e) {
		WebUI.enhancedClick(firstAgencyRow, FailureHandling.OPTIONAL)
	}

	WebUI.delay(1)
	waitBlockUI(20)
	WebUI.delay(1)

	// =========================
	// Click Select Button
	// =========================
	c(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Select Button'))
	waitBlockUI(20)
	WebUI.delay(0.5)
}

// =========================
// Get Contract No.
// =========================
TestObject contractNoObj = new TestObject('contractNoObj')
contractNoObj.addProperty(
	"xpath",
	ConditionType.EQUALS,
	"//td[label[normalize-space()='Contract No.']]/following-sibling::td[contains(@class,'header-info-text')][1]"
)

WebUI.waitForElementVisible(contractNoObj, 20)

String contractNo = WebUI.getText(contractNoObj).trim()
println("Contract No = " + contractNo)


// =========================
// Set File Ref Agency
// =========================
String fileRefAgencyInput = "File Ref " + contractNo
println("File Ref Agency = " + fileRefAgencyInput)

t(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/File Ref Agency'), fileRefAgencyInput)
WebUI.delay(0.5)

// =========================
// Upload Approval Letter Agency Attachment
// =========================
String uploadFilePath = System.getProperty("user.dir") + "/TestData/UploadFiles/File_pdf_for_testing.pdf"

c(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Upload Icon/Click Upload Button'))
waitBlockUI(20)
WebUI.delay(0.5)

up(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Upload Icon/Click Icon Choose File'), uploadFilePath,3)
waitBlockUI(20)
WebUI.delay(0.5)

c(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Upload Icon/Click Upload File Icon'))
waitBlockUI(20)
WebUI.delay(0.5)

c(findTestObject('Object Repository/FD and Agreement/FD Application/Agency/Upload Icon/Click Close button'))
waitBlockUI(20)
WebUI.delay(0.5)

// =========================
// Schedule
// =========================

TestObject scheduleMenu = findTestObject('Object Repository/FD and Agreement/Side Menu/FD Application/Schedule Side Menu')

if (WebUI.waitForElementClickable(scheduleMenu, 5, FailureHandling.OPTIONAL)) {

	c(scheduleMenu)
	waitBlockUI(20)
	WebUI.delay(0.5)

	int scheduleCount = 2

	for (int i = 0; i < scheduleCount; i++) {

		c(findTestObject('Object Repository/FD and Agreement/FD Application/Schedule/Schedule Add Button'))
		waitBlockUI(20)
		WebUI.delay(0.5)

		TestObject fromYear = new TestObject("fromYear_${i}")
		fromYear.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='_ctFulfilmentDetail_WAR_NGePportlet_:form:payScheduleTableId:${i}:fromYear_label']")
		selectDropdownByIndex(fromYear, Startyear)
		
		TestObject fromMonth = new TestObject("fromMonth_${i}")
		fromMonth.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='_ctFulfilmentDetail_WAR_NGePportlet_:form:payScheduleTableId:${i}:fromMonth_label']")
		selectDropdownByIndex(fromMonth, Startmonth)

		TestObject toYear = new TestObject("toYear_${i}")
		toYear.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='_ctFulfilmentDetail_WAR_NGePportlet_:form:payScheduleTableId:${i}:toYear_label']")
		selectDropdownByIndex(toYear, Endyear)

		TestObject toMonth = new TestObject("toMonth_${i}")
		toMonth.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='_ctFulfilmentDetail_WAR_NGePportlet_:form:payScheduleTableId:${i}:toMonth_label']")
		selectDropdownByIndex(toMonth, Endmonth)
	
	}

	} else {
		println("Schedule menu not available / not clickable, skip to next step")
	}
	
	TestObject Appmenu = findTestObject('Object Repository/FD and Agreement/Side Menu/FD Application/Approver Settings')
	
	if (WebUI.waitForElementClickable(Appmenu, 5, FailureHandling.OPTIONAL)) {
		WebUI.click(Appmenu)
		waitBlockUI(20)
	} else {
		println("Approver Settings menu not clickable / not available")
	}
	
	String approverName = ApproverName.toString().trim()
	TestObject approver = new TestObject('approver_dynamic')
	approver.addProperty(
		"xpath",
		ConditionType.EQUALS,
		"//li[contains(@class,'ui-picklist-item') and normalize-space()='${approverName}']"
	)
	
	WebUI.waitForElementClickable(approver, 20)
	WebUI.click(approver)
	WebUI.delay(0.5)
	
	c(findTestObject('Object Repository/FD and Agreement/FD Application/Approver Setting/Approver Right button'))
	waitBlockUI(10)
	WebUI.delay(0.5)
	
	c(findTestObject('Object Repository/FD and Agreement/Submit Button'))
	waitBlockUI(10)
	WebUI.delay(0.5)

	/* =========================
	 * SUCCESS MESSAGE - CT ONLY
	 * Purpose:
	 * - wait for loader disappear
	 * - capture success message
	 * - extract dynamic CT number
	 * ========================= */
	TestObject blockUI = new TestObject('blockUI')
	blockUI.addProperty("xpath", ConditionType.EQUALS,
		"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
	)
	
	if (WebUI.verifyElementPresent(blockUI, 2, FailureHandling.OPTIONAL)) {
		WebUI.waitForElementNotVisible(blockUI, 30, FailureHandling.OPTIONAL)
	}
	
	TestObject msgObj = new TestObject('msg_CT_saved')
	msgObj.addProperty("xpath", ConditionType.EQUALS,
		"//span[contains(@class,'ui-messages-info-detail') and " +
		"contains(.,'Fulfilment Details Creation') and " +
		"contains(.,'is successfully submitted to Contract Approver')]"
	)
	
	WebUI.waitForElementVisible(msgObj, 30)
	
	String msg = ""
	for (int i = 0; i < 2; i++) {
		msg = WebUI.getText(msgObj, FailureHandling.OPTIONAL)
		if (msg != null && msg.contains("CT")) break
		WebUI.delay(1)
	}
	
	msg = (msg == null) ? "" : msg.trim()
	WebUI.comment("Message: " + msg)
	
	def matcher = (msg =~ /(CT\d+)/)
	String ctNo = matcher.find() ? matcher.group(1) : ""
	
	if (ctNo == "") {
		WebUI.takeScreenshot()
		assert false : "❌ CT number not found. Message was: " + msg
	}
	
	WebUI.comment("✅ Captured CT No: " + ctNo)
	
	/* =========================
	 * EXCEL APPEND
	 * Purpose:
	 * - append CT number and message into same Excel file
	 * ========================= */
	String baseDir = System.getProperty("user.home") + "/Desktop/PrepDataFileNumber"
	new File(baseDir).mkdirs() //AUTO-CREATE FOLDER
	String filePath = baseDir + "/FD_Submission_AP_201_2026.xlsx"
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
		header.createCell(1).setCellValue("CT No")
		header.createCell(2).setCellValue("Message")
	}
	
	if (fis != null) fis.close()
	
	int nextRow = (sheet.getPhysicalNumberOfRows() == 0) ? 0 : sheet.getLastRowNum() + 1
	def row = sheet.createRow(nextRow)
	
	row.createCell(0).setCellValue(now)
	row.createCell(1).setCellValue(ctNo)
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
