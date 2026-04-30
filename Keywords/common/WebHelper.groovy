package common

import com.kms.katalon.core.testobject.*
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.webui.common.WebUiCommonHelper
import org.openqa.selenium.WebElement
import org.openqa.selenium.StaleElementReferenceException
import java.math.BigDecimal
import java.util.Arrays

class WebHelper {

	// Convert excel/csv value safely
	static int toInt(def v, int defaultVal = 0) {
		if (v == null) return defaultVal
		return new BigDecimal(v.toString().trim()).intValue()
	}

	// PrimeFaces overlay wait
	static def waitBlockUI(int timeout = 30) {
		TestObject blockUI = new TestObject('blockUI')
		blockUI.addProperty("xpath", ConditionType.EQUALS,
			"//*[contains(@class,'ui-blockui') or contains(@class,'blockUI') or contains(@class,'ui-widget-overlay')]"
		)

		if (WebUI.verifyElementPresent(blockUI, 1, FailureHandling.OPTIONAL)) {
			WebUI.waitForElementNotVisible(blockUI, timeout, FailureHandling.OPTIONAL)
		}
	}

	// wait visible
	static def wVisible(TestObject obj, int timeout = 1) {
		waitBlockUI(Math.min(timeout, 1))
		WebUI.waitForElementVisible(obj, timeout, FailureHandling.STOP_ON_FAILURE)
	}

	// wait clickable
	static def wClickable(TestObject obj, int timeout = 1) {
		wVisible(obj, timeout)
		WebUI.waitForElementClickable(obj, timeout, FailureHandling.STOP_ON_FAILURE)
	}

	// click
	static def c(TestObject obj, int timeout = 1) {
		for (int i = 0; i < 2; i++) {
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

		wClickable(obj, timeout)
		WebUI.click(obj)
		waitBlockUI(1)
	}

	// double click
	static def dc(TestObject obj, int timeout = 1) {
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

	// set text
	static def t(TestObject obj, def value, int timeout = 1) {
		wVisible(obj, timeout)
		WebUI.scrollToElement(obj, 1, FailureHandling.OPTIONAL)
		WebUI.setText(obj, (value == null ? "" : value.toString()))
	}

	// upload
	static def up(TestObject obj, String filePath, int timeout = 1) {
		wVisible(obj, timeout)
		WebUI.uploadFile(obj, filePath)
		waitBlockUI(1)
	}

	// set zone qty
	static def setZoneQtyByRow(int rowIndex, String qtyValue) {
		String xpath =
			"//div[contains(@class,'ui-dialog')]//input[contains(@id,'specZoneQtyTbl:${rowIndex}:zoneQty')]"

		TestObject qtyObj = new TestObject("zoneQty_" + rowIndex)
		qtyObj.addProperty("xpath", ConditionType.EQUALS, xpath)

		WebUI.waitForElementVisible(qtyObj, 20)
		WebElement qtyEl = WebUiCommonHelper.findWebElement(qtyObj, 20)

		WebUI.executeJavaScript("""
			arguments[0].value = arguments[1];
			arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
			arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
		""", Arrays.asList(qtyEl, qtyValue))

		waitBlockUI(30)
		WebUI.delay(0.5)
	}

	// set unit price
	static def setUnitPriceByRow(int rowIndex, String unitPriceValue) {
		String xpath =
			"//div[contains(@class,'ui-dialog')]//input[contains(@id,'specAnswerTbl:${rowIndex}:ratePerUomAns')]"

		TestObject priceObj = new TestObject("unitPrice_" + rowIndex)
		priceObj.addProperty("xpath", ConditionType.EQUALS, xpath)

		WebUI.waitForElementVisible(priceObj, 20)
		WebElement priceEl = WebUiCommonHelper.findWebElement(priceObj, 20)

		WebUI.executeJavaScript("""
			arguments[0].value = arguments[1];
			arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
			arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
		""", Arrays.asList(priceEl, unitPriceValue))

		waitBlockUI(30)
		WebUI.delay(0.5)
	}

	// open PrimeFaces dropdown
	static def openPFDropdown(TestObject triggerObj) {
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

	// click dropdown option
	static def clickPFOptionByIndex(int index0) {
		TestObject opt = new TestObject("pfOpt_" + index0)
		opt.addProperty("xpath", ConditionType.EQUALS,
			"(//div[contains(@class,'ui-selectonemenu-panel') and contains(@style,'display: block')]//li[contains(@class,'ui-selectonemenu-item')])[${index0 + 1}]"
		)

		c(opt, 20)
		WebUI.delay(0.2)
		waitBlockUI(20)
	}

	// universal dropdown select
	static def selectDropdownByIndex(TestObject dropdownObj, def indexFromData) {

		int idx0 = toInt(indexFromData)

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

			} catch (StaleElementReferenceException e) {
				WebUI.delay(0.5)
			}
		}

		assert false : "Dropdown failed: " + dropdownObj.getObjectId()
	}

	// claim document
	static boolean claimDocument(String targetDocNo) {

		boolean found = false

		for (int pageIndex = 1; pageIndex <= 10; pageIndex++) {

			TestObject targetRow = new TestObject("targetRow_" + targetDocNo)
			targetRow.addProperty("xpath", ConditionType.EQUALS,
				"//tbody[contains(@id,'taskListGroupId_data')]//tr[td[normalize-space()='" + targetDocNo + "']]"
			)

			if (WebUI.verifyElementPresent(targetRow, 3, FailureHandling.OPTIONAL)) {

				TestObject claimBtn = new TestObject("claimBtn_" + targetDocNo)
				claimBtn.addProperty("xpath", ConditionType.EQUALS,
					"//tbody[contains(@id,'taskListGroupId_data')]//tr[td[normalize-space()='" + targetDocNo + "']]//span[normalize-space()='Claim']/ancestor::button"
				)

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
}