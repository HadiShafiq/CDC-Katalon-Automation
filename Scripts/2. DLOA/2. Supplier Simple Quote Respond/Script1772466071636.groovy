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
 * DLOA - Simple Quote Respond
 * ========================= */

WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Click Simple Quote Respond'))

String targetSQ = sqNo
boolean found = false

for (int pageIndex = 1; pageIndex <= 10; pageIndex++) {

	WebUI.comment("Checking page index: " + pageIndex + " for SQ No: " + targetSQ)

	TestObject targetRow = new TestObject("targetRow_" + pageIndex)
	targetRow.addProperty("xpath", ConditionType.EQUALS,
		"//*[@id='_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId_data']//tr[contains(.,'${targetSQ}')]"
	)

	if (WebUI.verifyElementPresent(targetRow, 3, FailureHandling.OPTIONAL)) {
		WebUI.comment("✅ Found SQ Number: " + targetSQ + " on page index " + pageIndex)

		TestObject procurementTitle = new TestObject("procurementTitle_" + pageIndex)
		procurementTitle.addProperty("xpath", ConditionType.EQUALS,
			"(//*[@id='_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId_data']//tr[contains(.,'${targetSQ}')]//*[starts-with(@id,'_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId:')]/span)[1]"
		)

		c(procurementTitle, 20)
		waitBlockUI(20)

		found = true
		break
	}

	if (pageIndex < 10) {
		TestObject nextPage = new TestObject("page_" + (pageIndex + 1))
		nextPage.addProperty("xpath", ConditionType.EQUALS,
			"//*[@id='_scSupplierInvitationSq_WAR_NGePportlet_:form:scSuppInviteTblId_paginator_bottom']/span[4]/span[${pageIndex + 1}]"
		)

		if (WebUI.verifyElementClickable(nextPage, FailureHandling.OPTIONAL)) {
			WebUI.comment("SQ not found on page " + pageIndex + ". Moving to page " + (pageIndex + 1))
			c(nextPage, 20)
			waitBlockUI(20)
			WebUI.delay(1)
		}
	}
}

assert found : "❌ SQ Number not found from pagination 1 until 10: " + targetSQ

//WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Click Procument Tittle'))

// Fill in the first unit price field
t(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Unit Price(RM)'),
	UnitPrice, 20)
waitBlockUI(20)
WebUI.delay(3)

// Send Tab key to move to the next input field
WebUI.sendKeys(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Unit Price(RM)'), Keys.chord(Keys.TAB))

// Fill in the second unit price field
//t(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Unit Price(RM) - 2'),
	//UnitPrice2, 20)

// Send Tab key to move to the next field (or to the next available element)
//WebUI.sendKeys(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Unit Price(RM) - 2'), Keys.chord(Keys.TAB))
//waitBlockUI(20)
//WebUI.delay(3)

WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Submit Respond Simple Quote'))
waitBlockUI(20)
WebUI.delay(3)

WebUI.click(findTestObject('Object Repository/DLOA/5. SUpplier Simple Quote Respond/Click Soft Cert Sign Button'))
waitBlockUI(20)
WebUI.delay(3)

/* =========================
 * Sign Out
 * ========================= */
WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Menu For Sign Out'))

WebUI.click(findTestObject('Object Repository/Direct LOA/1. Direct LOA Requistioner/LogOut/Click Sign Out'))

// wait until logout is completed (choose one)
WebUI.waitForPageLoad(20)
WebUI.closeBrowser()


