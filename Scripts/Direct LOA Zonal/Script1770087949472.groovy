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

WebUI.selectOptionByValue(findTestObject('null'), 'en_US', true)

WebUI.click(findTestObject('null'))

WebUI.setText(findTestObject('null'), '710527085562')

WebUI.setText(findTestObject('null'), 'P@ssw0rd1234')

WebUI.click(findTestObject('null'))

WebUI.click(findTestObject('null'))

WebUI.click(findTestObject('Object Repository/Zone item_Zonal/Page_Task List - ePerolehan/span_My Task_ui-icon ui-icon-plusthick'))

WebUI.setText(findTestObject('Object Repository/Zone item_Zonal/Page_Task List - ePerolehan/input_Document No__BpmTaskList_WAR_NGePport_a136e8'), 
    'LA260000000000460')

WebUI.click(findTestObject('Object Repository/Zone item_Zonal/Page_Task List - ePerolehan/span_Search'))

WebUI.click(findTestObject('null'))

//click side menu Zone Item
WebUI.click(findTestObject('Object Repository/Zone item_Zonal/Side Menu/Side Menu Zone Item'))

// Zonal
WebUI.click(findTestObject('null'))
 
 WebUI.click(findTestObject('null'))
 
 WebUI.click(findTestObject('null'))
 
 WebUI.click(findTestObject('null'))
 
 WebUI.click(findTestObject('null'))
 
 WebUI.setText(findTestObject('null'),
	 'Zonal')
 
 WebUI.click(findTestObject('null'))
 
 //Product
 WebUI.click(findTestObject('null'))
 
 WebUI.setText(findTestObject('null'),
	 'Testing')
 
 TestObject uom = findTestObject('null')
 
 WebUI.click(uom)
 
 WebUI.setText(uom, 'box')
 
 WebUI.delay(2 // bagi result list muncul
	 )
 
 WebUI.sendKeys(uom, Keys.chord(Keys.ENTER))
 
 WebUI.selectOptionByValue(findTestObject('null'),
	 '298', true)
 
 // kne add click pencil icon for quantity
 WebUI.click(findTestObject('Object Repository/Zone item_Zonal/Zonal/Object Repository/Zone item/Page_Letter Of Acceptance (LOA) - ePerolehan/Zonal/span_Grand Total (RM)_ui-button-icon-left u_2834a0'))
 
 
 WebUI.setText(findTestObject('null'),
	 '10')
 
 WebUI.click(findTestObject('null'))
 
 WebUI.click(findTestObject('null'))
 
 WebUI.setText(findTestObject('null'),
	 '30000')
 
 WebUI.click(findTestObject('null'))
 
 WebUI.click(findTestObject('null'))

WebUI.click(findTestObject('null'))

WebUI.setText(findTestObject('null'),
	'Product')
 
 
/* 
*/
 

/*




WebUI.setText(findTestObject('null'), 
    '10')

WebUI.setText(findTestObject('null'), 
    '300000')


//Service
// Dynamic TestObject for Click Button Add SERVICE
TestObject AddService = findTestObject('null')

scrollToElement(AddService, 5)

waitForElementClickable(AddService, 20)

click(AddService)

//Add Specification
WebUI.setText(findTestObject('null'), 
    'testing')

//Add OUM on service table
TestObject uomService = findTestObject('null')

WebUI.waitForElementVisible(uomService, 20)

WebUI.click(uomService)

WebUI.setText(uomService, 'box')

WebUI.delay(2)

WebUI.sendKeys(uomService, Keys.chord(Keys.ENTER))

//Add Freq. Per UOM
WebUI.setText(findTestObject('null'), 
    '10')

WebUI.setText(findTestObject('null'), 
    '300000')

WebUI.setText(findTestObject('null'), 
    '10')

WebUI.click(findTestObject('null'))

// Click SERVICE Specification (last visible)
def spec = findTestObject('null')

WebUI.waitForElementVisible(spec, 10)

def el = WebUiCommonHelper.findWebElement(spec, 10)

DriverFactory.getWebDriver().executeScript('arguments[0].click();', el)

WebUI.setText(findTestObject('null'), 
    'service')

WebUI.click(findTestObject('null'))*/




