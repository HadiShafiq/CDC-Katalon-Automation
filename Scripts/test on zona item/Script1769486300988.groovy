import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.testobject.TestObject

import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testobject.ConditionType as ConditionType
import com.kms.katalon.core.webui.common.WebUiCommonHelper as WebUiCommonHelper
import com.kms.katalon.core.webui.driver.DriverFactory as DriverFactory
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import org.openqa.selenium.JavascriptExecutor as JavascriptExecutor
import org.openqa.selenium.WebDriver as WebDriver
import org.openqa.selenium.chrome.ChromeDriver as ChromeDriver
import org.openqa.selenium.chrome.ChromeOptions as ChromeOptions
import java.nio.file.Files as Files
import org.openqa.selenium.Keys as Keys
import com.kms.katalon.core.testobject.ConditionType
import static com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords.*

/* =========================
 * PrimeFaces Helpers (keep simple)
 * ========================= */
// open PrimeFaces dropdown (sometimes need 2 clicks)
def openPFDropdown(TestObject triggerObj) {

	TestObject panelOpen = new TestObject('pfPanelOpen')
	panelOpen.addProperty("xpath", ConditionType.EQUALS,
		"//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]"
	)

	WebUI.click(triggerObj)
	WebUI.delay(0.3)

	if (!WebUI.waitForElementVisible(panelOpen, 1, FailureHandling.OPTIONAL)) {
		WebUI.click(triggerObj)
		WebUI.delay(0.3)
		WebUI.waitForElementVisible(panelOpen, 3, FailureHandling.OPTIONAL)
	}
}

// click dropdown option by index (0-based)
def clickPFOptionByIndex(int index) {
	TestObject opt = new TestObject("pfOpt_" + index)
	opt.addProperty("xpath", ConditionType.EQUALS,
		"(//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[contains(@class,'ui-selectonemenu-item')])[${index + 1}]"
	)

	WebUI.waitForElementClickable(opt, 5)
	WebUI.click(opt)
	WebUI.delay(0.2)
}
	
/* =========================
 * Browser Setup (Guest + Clean)
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

options.setExperimentalOption('prefs', [('credentials_enable_service') : false, ('profile.password_manager_enabled') : false
        , ('profile.default_content_setting_values.notifications') : 2])

WebDriver driver = new ChromeDriver(options)

DriverFactory.changeWebDriver(driver)
/* =========================
 * Test Flow
 * ========================= */
WebUI.navigateToUrl('http://ngepsit.eperolehan.com.my/home')

WebUI.maximizeWindow()

WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Dropdown Language'), 'en_US', true)

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Login/Right Top Menu Login'))

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Login/Username'), '710527085562')

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Login/Password'), 'P@ssw0rd1234')

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Login/Submit Username and Password'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Common Page/Click Task List'))

WebUI.click(findTestObject('Object Repository/Zone item_Zonal/Page_Task List - ePerolehan/span_My Task_ui-icon ui-icon-plusthick'))

WebUI.setText(findTestObject('Object Repository/Zone item_Zonal/Page_Task List - ePerolehan/input_Document No__BpmTaskList_WAR_NGePport_a136e8'), 
    'LA260000000000460')

WebUI.click(findTestObject('Object Repository/Zone item_Zonal/Page_Task List - ePerolehan/span_Search'))

WebUI.click(findTestObject('Object Repository/Direct LOA/2. Direct LOA Supplier/TaskList Supplier/Click TaskList Description'))


//click side menu Zone Item 
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Zone Item'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Click Action Item for Product'))

//Product

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input Specification 1 TextBox'),
	'Testing')

TestObject uom = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input UOM')

WebUI.click(uom)
WebUI.setText(uom, 'box')
WebUI.delay(2)  // bagi result list muncul
WebUI.sendKeys(uom, Keys.chord(Keys.ENTER))

WebUI.selectOptionByIndex(
	findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Dropdown Price Type'),
	1
)

//WebUI.selectOptionByValue(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Dropdown Price Type'),
	//'294', true)

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Quantity'),
	'10')
WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Unit Price(RM)'),
	'300000')

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Action Item Click Specification'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Click Specification'))

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Product/Input Specification 2 TextBox'),
	'Product')

//Service


// Dynamic TestObject for Click Button Add SERVICE


TestObject AddService = findTestObject ('Zone item/Page_Letter Of Acceptance (LOA) - ePerolehan/Service/Click Action Item for Service')

scrollToElement(AddService, 5)
waitForElementClickable(AddService, 20)
click(AddService)

//Add Specification
WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Input Specification 1 TextBox'),
	'testing')

//Add OUM on service table

TestObject uomService = findTestObject('Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Input UOM')

WebUI.waitForElementVisible(uomService, 20)
WebUI.click(uomService)
WebUI.setText(uomService, 'box')
WebUI.delay(2)

WebUI.sendKeys(uomService, Keys.chord(Keys.ENTER))

//Add Freq. Per UOM

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Quantity'),
	'10')

TestObject UnitPrice = findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Unit Price(RM)')

WebUI.setText(UnitPrice, '300000')

WebUI.sendKeys(UnitPrice, Keys.chord(Keys.TAB))
WebUI.delay(3)
//WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Unit Price(RM)'),'300000')

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Freq per UOM'),
	'10')

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Action Item Click Specification'))

// Click SERVICE Specification (last visible)
def spec = findTestObject('Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Click Specification')
WebUI.waitForElementVisible(spec, 10)

def el = WebUiCommonHelper.findWebElement(spec, 10)
DriverFactory.getWebDriver().executeScript("arguments[0].click();", el)

WebUI.setText(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Zone Item Tab/Add Service/Input Specification 2 TextBox'),
	'service')

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/Side Menu/Side Menu Payment Deduction'))