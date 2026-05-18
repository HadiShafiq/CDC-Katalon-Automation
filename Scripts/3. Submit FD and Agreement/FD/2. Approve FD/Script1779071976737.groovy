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

/* ==================================================
 * HELPERS -- Dalam folder Keywords/common/WebHelper.groovy
 * Purpose:
 * - reusable utility for conversion, wait, click, text input, upload
 * - keep script stable without changing main flow
 * ================================================== */
import static common.WebHelper.* 

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

//wait loader gone
waitBlockUI(30)

//wait table/data loaded (VERY IMPORTANT)
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

//click Approve Button 
c(findTestObject('Object Repository/FD and Agreement/FD Application/FD Approver/Approve Button'))
waitBlockUI(20)
WebUI.delay(0.5)

//click Sign
c(findTestObject('Object Repository/FD and Agreement/FD Application/FD Approver/Click Sign'))
waitBlockUI(20)
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
	"contains(.,'is approved')]"
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
String filePath = baseDir + "/FD_Approver_201_2026.xlsx"
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