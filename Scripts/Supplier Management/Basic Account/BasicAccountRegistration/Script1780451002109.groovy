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
 * Click Account type button by option number.
 * 1 = Basic
 * 2 = MOF
 * 3 = G2G
 */
def clickAccountType(int option) {

	String xpath =
		"(//div[contains(@class,'ui-radiobutton')]//input[@type='radio'])[${option}]"

	TestObject obj = new TestObject("accountType_" + option)
	obj.addProperty("xpath", ConditionType.EQUALS, xpath)

	WebUI.waitForElementPresent(obj, 10)
	WebUI.scrollToElement(obj, 5)

	WebElement element = WebUiCommonHelper.findWebElement(obj, 10)

	WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(element))

	waitBlockUI(10)
	WebUI.delay(0.3)
}

/**
 * For Bumiputera or non Bumi
 * Yes = Bumiputera
 * No = non Bumiputera
 */

def clickBumiputera(String option) {

    String boxXpath

    if (option.equalsIgnoreCase("Yes")) {
        boxXpath = "//table[contains(@id,'bumiId')]//label[normalize-space()='Yes']/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
    } else if (option.equalsIgnoreCase("No")) {
        boxXpath = "//table[contains(@id,'bumiId')]//label[normalize-space()='No']/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
    } else {
        throw new Exception("Invalid option: " + option)
    }

    TestObject obj = new TestObject("bumi_" + option)
    obj.addProperty("xpath", ConditionType.EQUALS, boxXpath)

    WebUI.waitForElementClickable(obj, 10)
    WebUI.scrollToElement(obj, 5)

    WebElement element = WebUiCommonHelper.findWebElement(obj, 10)
    WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(element))

    waitBlockUI(10)
    WebUI.delay(0.3)
}
/**
 * For Muslim or non Muslim
 * Yes = Muslim
 * No = non Muslim
 */
def clickReligion(int religion) {

    String boxXpath

    switch (religion) {
        case 1: // Muslim
            boxXpath = "//table[contains(@id,'religionId')]//label[normalize-space()='Muslim']/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
            break

        case 2: // Non - Muslim
            boxXpath = "//table[contains(@id,'religionId')]//label[contains(normalize-space(),'Non')]/preceding::div[contains(@class,'ui-radiobutton-box')][1]"
            break

        default:
            throw new Exception("Invalid religion value: ${religion}. Use 1=Muslim, 2=Non Muslim.")
    }

    TestObject obj = new TestObject("religion_" + religion)
    obj.addProperty("xpath", ConditionType.EQUALS, boxXpath)

    WebUI.waitForElementClickable(obj, 10)
    WebUI.scrollToElement(obj, 5)

    WebElement element = WebUiCommonHelper.findWebElement(obj, 10)

    WebUI.executeJavaScript("arguments[0].click();", Arrays.asList(element))

    waitBlockUI(10)
    WebUI.delay(0.3)
}
/* =========================================================
 * 7) CALENDAR PICKER DATE - Supplier Registration Page only
 * ========================================================= */

def pickDate(String yyyyMmDd) {

    def parts = yyyyMmDd.split('-')
    int targetYear  = parts[0] as int
    int targetMonth = parts[1] as int
    String targetDay = (parts[2] as int).toString()

    TestObject monthObj = new TestObject()
    monthObj.addProperty("xpath", ConditionType.EQUALS,
        "//*[@id='ui-datepicker-div']//select[contains(@class,'ui-datepicker-month')]")

    TestObject yearObj = new TestObject()
    yearObj.addProperty("xpath", ConditionType.EQUALS,
        "//*[@id='ui-datepicker-div']//select[contains(@class,'ui-datepicker-year')]")

    WebUI.waitForElementVisible(monthObj, 10)

    WebUI.selectOptionByValue(monthObj, String.valueOf(targetMonth - 1), false)
    WebUI.selectOptionByValue(yearObj, String.valueOf(targetYear), false)

    String dayXpath =
        "//*[@id='ui-datepicker-div']//td[not(contains(@class,'ui-datepicker-other-month')) " +
        "and not(contains(@class,'ui-state-disabled'))]//a[normalize-space()='" + targetDay + "']"

    TestObject dayObj = new TestObject()
    dayObj.addProperty("xpath", ConditionType.EQUALS, dayXpath)

    WebUI.waitForElementClickable(dayObj, 10)
    WebUI.click(dayObj)
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
wVisible(findTestObject('Object Repository/Home Page/LanguageSelect'), 20)
WebUI.selectOptionByValue(findTestObject('Object Repository/Home Page/LanguageSelect'), 'en_US', true)
	waitBlockUI(20)
	WebUI.delay(0.5)
	WebUI.delay(1)

//Click New Supplier Registration
c(findTestObject('Object Repository/Home Page/NewSupplierRegistration'))
	waitBlockUI(20)
	
//Click Basic Account radio button
//int rbAccountType declare on variable
clickAccountType(rbAccountType)
	
//Next
c(findTestObject('Object Repository/SupplierRegistration/AccTypePage/AccTypeNext'))
	waitBlockUI(20)

//Captcha page
c(findTestObject('Object Repository/SupplierRegistration/CaptchaPage/TnCBox'))
waitBlockUI(20)
	c(findTestObject('Object Repository/SupplierRegistration/CaptchaPage/CaptchaNext'))
	waitBlockUI(20)

//Account information page

// Handle different menu based on business type
	def handleBusinessMenu(int type) {
		
			switch(type) {
		
				case 1:
				case 2:
					handleMenuGroupA() //Sabah Partnership/Sole-Propriertorship
					break
		
				case 3:
				case 4:
					handleMenuGroupB()
					break
		
				case 5:
				case 6:
					handleMenuGroupC()
					break
		
				case 7:
				case 8:
					handleMenuGroupD()
					break
				case 9:
					handleMenuGroupE()
					break
		
				case 10:
				case 11:
				case 12:
				case 13:
					handleMenuGroupF()
					break
		
				case 14:
					handleMenuGroupG()
					break
		
				default:
					KeywordUtil.markFailed("Unsupported BusinessCompanyType: ${type}")
			}
		}

//Each menu group
		
	// menu logic for types 1,2 Sabah Partnership/Sole Propriertorship
	def handleMenuGroupA() {
		//State selection - state must be selected before local authority	
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessState'), State)
		//District
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessDistrict'), District)
		//City/Town
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessCity'), City)
		waitBlockUI(20)
		//Local authority select
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/LocalAuthority'), LocalAuthority)
		waitBlockUI(20)
	
		
	}
	// Sarawak 3,4
	def handleMenuGroupB() {
		
		//State selection - state must be selected before local authority
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessState'), State)
		//Division
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/CompanyPhone/Division'), Division)
		//District
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessDistrict'), District)
		//City/Town
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessCity'), City)
		waitBlockUI(20)
		//Local authority select
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/LocalAuthority'), LocalAuthority)
		waitBlockUI(20)
		
	}
	//ROB Partnership and ROB sole proprietorship 5,6
	def handleMenuGroupC() {
		//State
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessState'), State)
		//District
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessDistrict'), District)
		//City/Town
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessCity'), City)
		}
		
	//ROC 7,8
	def handleMenuGroupD() {
		//State
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessState'), State)
		//District
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessDistrict'), District)
		//City/Town
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessCity'), City)
		//Owner Information
			//Tick Equity Owner
			//t(findTestObject(''))
	}
	
	// LLP 9
	def handleMenuGroupE() {
		//State
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessState'), State)
		//District
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessDistrict'), District)
		//City/Town
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessCity'), City)
	}
	//Cooperative 10, Society 11, ROS Organization 12, Organization - others 13
	def handleMenuGroupF() {
		//State
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessState'), State)
		//District
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessDistrict'), District)
		//City/Town
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessCity'), City)
	}
	
	//Individual 14
	def handleMenuGroupG() {
		
		//Service 
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/IndividualRegistration/ServiceDropdown'), IndividualService)
		//Bumiputera and Religion radio
		clickBumiputera(Bumiputera)
			clickReligion(Religion)
			
		//Date of Birth
		c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/DateOfBirth-Calendar'))
			WebUI.delay(1)
				pickDate("2003-05-11")
		//Owner IC
		t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/SuppRegIC'), OwnerIC)
		waitBlockUI(20)
		//Owner Name
		t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/IndividualRegistration/IndividualName'), OwnerName)
		waitBlockUI(20)
		//Address	
		t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminAddress'),AdminAddress)
			waitBlockUI(10)
			//State
			selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/State'),SupplierState)
			
			//District //District depends on state
			selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/District'),SupplierDistrict)
			
			// City/Town
			selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/City-Town'),SupplierCity)
		
			//Admin Postcode + type slow
			TestObject adminPostcodeField = findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminPostcode')
			String postcodeValue = AdminPostcode.toString()
			
			WebUI.click(adminPostcodeField)
			
			for (char c : postcodeValue.toCharArray()) {
				WebUI.sendKeys(adminPostcodeField, c.toString())
				WebUI.delay(0.2)
			}
			//Bank Information
			c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankInfoAddBtn'))
			waitBlockUI(20)
			//Bank name dropdown
			selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankName'), BankName)
			waitBlockUI(20)
			//Bank acc number
			WebUI.waitForElementVisible(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankAccountNumber'),10)
			t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankAccountNumber'), BankAccNum)
			
			
		//Email and confirm email
		t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Email'),SupplierEmail)
		// Confirm
		t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Email-Confirmation'),SupplierEmail)
		
		
		//Admin Phone Number
		String AdminCountryCode = AdminNum.substring(0, 3)
		String AdminMobileCode = AdminNum.substring(3, 4)
		String AdminMobile = AdminNum.substring(4)
			t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminMobileCode'), AdminMobileCode)
			waitBlockUI(1)
			t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminMobile'), AdminMobile)
			waitBlockUI(1)
			
	//Mobile Telco
		selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminMobileTelco'),'1')
	//Correspondence 
	c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/CorrespondenceAddress'))
		waitBlockUI(10)
		
	//Submit and accept terms
c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Submission/SubmitSupplierRegistration'))
	waitBlockUI(10)
c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Submission/I_AcceptTerms'))
	waitBlockUI(10)
c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Submission/Submission_Proceed'))
	waitBlockUI(10)
	}	//End of Individual 
		
			
		
//Account Information box
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/AccInfo/Bus_CompType'), BusinessCompanyType)
	waitBlockUI(20)
	
	//Business Registration number
	t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/AccInfo/Bus_RegNo'), BusinessRegistrationNo)
	
//when SSM on, this must be on
	//c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/AccInfoPageNext'))
	//waitBlockUI(20)
	
	
	//End SSM section
	
	//Choose Menu based on business type
	handleBusinessMenu(BusinessCompanyType)
	
	// Stop script for Individual 
	if (BusinessCompanyType == 14) {
		return
	}

//Company Information box
	
	// Business/Company Name
	t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/BusinessName'), BusinessName)
	
	//Calendar
	// Click date input/calendar icon first
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/CompanyInfoCalendar'))
	WebUI.delay(1) 
	pickDate("2026-05-03")
	
	//Company address
	t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/CompAddress1'), CompanyAddress)
	waitBlockUI(20)
		//Postcode
		t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/Postcode'), CompanyPostcode)
		waitBlockUI(20)
		
	//Map and Coordinate
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/ViewMap'))
	waitBlockUI(20)
		c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/CloseMap'))
		
	//Company Phone number
		String countryCode = CompanyNum.substring(0, 3)
		String areaCode = CompanyNum.substring(3, 4)
		String number = CompanyNum.substring(4)
			t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/CompanyPhone/CompAreaCode'), areaCode)
			waitBlockUI(20)
			t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/CompanyInfo/CompanyPhone/CompNumber'), number)
			waitBlockUI(20)

//Owner Information
	//Owner Name
	t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/OwnerInfo/OwnerName'), OwnerName)
	waitBlockUI(20)
	//Owner IC
	t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/OwnerInfo/OwnerIC'), OwnerIC)
	waitBlockUI(20)
	
//Nature of Business
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/NatureOfBusiness/NOB_Add'))
	waitBlockUI(20)
	//Select search, code = 1 description = 2
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/NatureOfBusiness/NOB_type'), '1')
	waitBlockUI(20)
	//Fill in the NOB code
	t(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/NatureOfBusiness/NOBCode'), NOBCode)
	waitBlockUI(20)
	
	//Click search
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/NatureOfBusiness/NOB_Search'))
	waitBlockUI(20)
	//Click 1st result //xpath
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/NatureOfBusiness/NOB_choice'))
	waitBlockUI(20)
	//Select
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/NatureOfBusiness/NOB_select'))
	waitBlockUI(20)
	
//Bank Information
	c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankInfoAddBtn'))
	waitBlockUI(20)
	//Bank name dropdown
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankName'), BankName)
	waitBlockUI(20)
	//Bank acc number + type slower
	TestObject bankField = findTestObject(
    'Object Repository/SupplierRegistration/AccInfoPage/BankInfo/BankAccountNumber')

String bankValue = BankAccNum.toString()

WebUI.click(bankField)

for (char c : bankValue.toCharArray()) {
    WebUI.sendKeys(bankField, c.toString())
    WebUI.delay(0.2)
}
//Bottom Next
		c(findTestObject('Object Repository/SupplierRegistration/AccInfoPage/AccInfoPageNext'))
		waitBlockUI(20)
		
		
// SUPPLIER ADMINSTRATOR INFORMATION PAGE *//

// Clicking both checkbox 		
	//Please click if supplier administrator same as owner in Company Information checkbox
		//Sometimes it doesn't click 
	WebUI.waitForElementClickable(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminSameAsOwner'), 10)
	c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminSameAsOwner'))
	waitBlockUI(10)
	
	// Correspondence Address
	c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/CorrespondenceAddress'))
	waitBlockUI(10)
	
//Salutation
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Salutation'),Salutation)

//Bumiputera radio button
	clickBumiputera(Bumiputera)
	waitBlockUI(10)
	
//Admin Address
	t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminAddress'),AdminAddress)
	waitBlockUI(10)
	

//Position in Company
	t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/PositionInCompany'),'Placeholder')
	waitBlockUI(10)
//Email and confirm email 
	t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Email'),SupplierEmail)
	// Confirm
	t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Email-Confirmation'),SupplierEmail)

//Date of Birth	
	c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/DateOfBirth-Calendar'))
	WebUI.delay(1)
		pickDate("2003-05-11") 
		
//Religion radio button
	clickReligion(Religion)

// 4 dropdowns	
	// Country let default to Malaysia

	//State 
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/State'),SupplierState)
	
	//District //District depends on state 
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/District'),SupplierDistrict)
	
	// City/Town
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/City-Town'),SupplierCity)

	//Admin Postcode + type slow
	TestObject adminPostcodeField = findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminPostcode')

	String postcodeValue = AdminPostcode.toString()
	
	WebUI.click(adminPostcodeField)
	
	for (char c : postcodeValue.toCharArray()) {
		WebUI.sendKeys(adminPostcodeField, c.toString())
		WebUI.delay(0.2)
	}
	
//Role in Company
	//Tick management
	c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/ManagementCheckbox'))
	waitBlockUI(1)
	
	//Tick staff
	c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/StaffCheckbox'))
	waitBlockUI(1)
	
	
	
//Admin Phone Number
	String AdminCountryCode = AdminNum.substring(0, 3)
	String AdminMobileCode = AdminNum.substring(3, 4)
	String AdminMobile = AdminNum.substring(4)
		t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminMobileCode'), AdminMobileCode)
		waitBlockUI(1)
		t(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminMobile'), AdminMobile)
		waitBlockUI(1)
		
//Mobile Telco 
	selectDropdownByIndex(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/AdminMobileTelco'),'1')

//Submit and accept terms
c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Submission/SubmitSupplierRegistration'))
	waitBlockUI(10)
c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Submission/I_AcceptTerms'))
	waitBlockUI(10)
c(findTestObject('Object Repository/SupplierRegistration/SupplierAdminInfoPage/Submission/Submission_Proceed'))
	waitBlockUI(10)
		