
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util.Collections;
using NUnit.Framework;
using NUnit.Framework.Internal;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

using NPOI.XSSF.UserModel.Helpers;
using ExcelDataReader;
using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System.Globalization;

namespace POM.CommonMethods
{
    internal class Base
    {
        public IWebDriver driver;
        public static long PAGE_LOAD_TIMEOUT = 20;
        public static long IMPLICIT_WAIT = 20;
        public static string TESTDATA_SHEET_PATH = "xlsx";







        public Base(IWebDriver driver)
        {
            PageFactory.InitElements(driver, this);
            this.driver = driver;

        }

        public void getUrl(string url)
        {

            driver.Url = url;

        }

        // *********************************HANDLE_IFRAME**********************************************
        public void switchToElement(IWebElement element)
        {
            driver.SwitchTo().Frame(element);

        }
        public void switchToFrameByName(string Name)
        {
            driver.SwitchTo().Frame(Name);

        }

        public void switchToFrameById(string Id)
        {
            driver.SwitchTo().Frame(Id);

        }
        public void switchToDefaultFrame()
        {
            driver.SwitchTo().DefaultContent();

        }
        // *********************************HIGHLIGHT_ELEMENT**********************************************
        public void highlightElement(IWebElement element)
        {
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript("arguments[0].style.border='2px solid red'", element);

        }
        // *********************************LOWLIGHT_ELEMENT**********************************************
        public void unhighlightElement(IWebElement element)
        {
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript("arguments[0].removeAttribute('style','')", element);
        }
        // *********************************DRAG_AND_DROP**********************************************
        public void dragAndDrop(IWebElement source, IWebElement target)
        {
            Actions actions = new Actions(driver);
            actions.DragAndDrop(source, target).Perform();

        }
        public void scrollUpPage()
        {
            Actions actions = new Actions(driver);
            actions.SendKeys(Keys.Up).Build().Perform();

        }
        // *********************************GET_ELEMENT_BY_LOCATOR**********************************************
        public IWebElement getElement(By locater)
        {
            IWebElement element = driver.FindElement(locater);
            return element;
        }
        // ***************************************CLICK********************************************************

        public void Click(IWebElement element)
        {
            element.Click();
        }
        // ********************************SENDKEYS*******************************************

        public void Sendkeys(IWebElement element, string input)
        {
            element.SendKeys(input);
        }

        //***************************************RIGHT_CLICK************************************************************
        public void RightClick(IWebElement element)
        {
            Actions action = new Actions(driver);
            action.ContextClick(element).Perform();
        }

        // ***************************************DOUBLE_CLICK************************************************************
        public void DoubleClick(IWebElement element)
        {
            Actions action = new Actions(driver);
            action.DoubleClick(element).Perform();
        }
        // ********************************CLEAR***********************************************

        public void Clear(IWebElement element)
        {
            element.Clear();
        }
        // ***************************************SENDKEYS_WITH_ACTION************************************************************
        public void doSendKeysWithAction(IWebElement element, string value)
        {
            Actions act = new Actions(driver);
            act.SendKeys(element, value).Perform();
        }
        // ***************************************CLICKANDHOLD_WITH_ACTION************************************************************
        public void clickAndHoldWithAction()
        {
            Actions act = new Actions(driver);
            act.ClickAndHold();
        }

        // ***************************************CLICK_LOCATOR_WITH_ACTION************************************************************
        public void doClickWithAction(IWebElement element)
        {
            Actions act = new Actions(driver);
            act.Click(element).Perform();
        }

        // ***************************************GET_TEXT************************************************************
        public string getText(IWebElement element)
        {
            string text = element.Text;
            return text;
        }
        // ***************************************ELEMENT_IS_DISPLAYED************************************************************
        public bool isDisplayed(IWebElement element)
        {
            bool value = element.Displayed;
            return value;

        }
        // ***************************************ELEMENT_IS_SELECTED*****************************************************
        public bool isSelect(IWebElement element)
        {
            bool value = element.Selected;
            return value;
        }
        //// ***************************************ELEMENT_IS_ENABLED*****************************************************
        public bool isEnable(IWebElement element)
        {
            return element.Enabled;
        }
        // ***********************************SELECT_FROM_DROPDOWN***********************************************
        // select by visible text
        public void SelectByVisibleText(IWebElement element, string text)
        {
            SelectElement select = new SelectElement(element);
            select.SelectByText(text);

        }


        // select by index
        public void SelectByIndex(IWebElement element, int index)
        {
            SelectElement select = new SelectElement(element);
            select.SelectByIndex(index);
        }
        // select by value
        public void SelectByValue(IWebElement element, string value)
        {
            SelectElement select = new SelectElement(element);
            select.SelectByValue(value);
        }

        // deselect by value
        public void deselectByValue(IWebElement element, string value)
        {
            SelectElement select = new SelectElement(element);
            select.DeselectByValue(value);

        }

        // deselectAll drop down
        public void deSelectAll(IWebElement element)
        {
            SelectElement select = new SelectElement(element);
            select.DeselectAll();
        }

        // deselectByVisibleText
        public void deselectByVisibleText(IWebElement element, string value)
        {
            SelectElement select = new SelectElement(element);
            select.DeselectByText(value);
        }
        // ***********************************CHECK_IF_SUPPORT_MULTISELECT_DROPDOWN***********************************************
        public bool isSelectMultiple(IWebElement element)
        {
            SelectElement select = new SelectElement(element);
            return select.IsMultiple;

        }
        // ***********************************CHECK_DEFAULT_VALUE_DROPDOWN***********************************************
        public string getDefaultSelectedValue(IWebElement element)
        {
            SelectElement select = new SelectElement(element);
            return select.WrappedElement.Text;

        }
        public ArrayList getOptions(IWebElement element)
        {
            SelectElement select = new SelectElement(element);
            var araylist = new ArrayList();
            for (int i = 0; i < select.Options.Count; i++)
            {
                araylist.Add(select.Options[i].Text);


            }
            return araylist;

        }

        // ***********************************SWITCH_WINDOW***********************************************
        public void SwitchWindows(IWebDriver driver)
        {
            string parent = driver.CurrentWindowHandle;
            IReadOnlyCollection<string> windowHandles = driver.WindowHandles;

            foreach (string childWindow in windowHandles)
            {
                if (parent != childWindow)
                {
                    driver.SwitchTo().Window(childWindow);
                    string childTitle = driver.Title;
                    Console.WriteLine(childTitle);

                    driver.Close();
                }
            }
            // again switching parent window
            driver.SwitchTo().Window(parent);
        }
        // *********************************SWITCH_TO_ACTIVE_WINDOW_TAB**********************************************
        public void SwitchToActiveWindow(IWebDriver driver)
        {
            List<string> windowHandles = new List<string>(driver.WindowHandles);

            // Switch to active tab (assuming the active window is at index 1)
            driver.SwitchTo().Window(windowHandles[1]);
        }
        // *********************************SWITCH_TO_OEN_WINDOW_TAB********************************
        public void SwitchToOpenWindow(IWebDriver driver, int id)
        {
            IList<string> windowHandles = new List<string>(driver.WindowHandles);

            // Switch to the window with the specified id
            driver.SwitchTo().Window(windowHandles[id]);
        }
        // *********************************MOUSE_HOVER**************
        public void MouseHover(IWebElement element)
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(element).Build().Perform();
        }
        // get Title
        public string GetTitle()
        {
            return driver.Title;
        }
        //get Text
        public string GetText(IWebElement element)
        {
            return element.Text;
        }
        // *************************************ALERT***************
        // To click on the ‘Cancel’ button of the alert.

        public void DismissAlert()
        {
            IAlert alert = driver.SwitchTo().Alert();
            alert.Dismiss();
        }
        // To click on the ‘OK’ button of the alert.
        public void AcceptAlert()
        {
            IAlert alert = driver.SwitchTo().Alert();
            alert.Accept();
        }
        //// To capture the alert message.
        public void GetAlertMessage()
        {
            IAlert alert = driver.SwitchTo().Alert();
            Console.WriteLine(alert.Text);
        }
        // To send some data to alert box.
        public void SendDataToAlertBox(string Text)
        {
            IAlert alert = driver.SwitchTo().Alert();
            alert.SendKeys(Text);
        }
        //// **********************************CHECKBOX*********
        public static void SelectCheckbox(IWebElement element)
        {
            element.Click();
        }
        //// multiple checkboxes
        public void SelectMultipleCheckboxes(List<IWebElement> listOfItems, int noOfCheckboxesToBeSelected)
        {
            for (int i = 0; i <= noOfCheckboxesToBeSelected - 1; i++)
            {
                listOfItems[i].Click();
                Console.WriteLine("Checkbox_" + i + " selected");
            }
        }
        // ****************************************UPLOAD_FILE**********************************

        public void UploadFile(IWebElement element, string filePath)
        {
            element.SendKeys(filePath);
        }
        // **********************************TAB******
        public void tAB(IWebElement element)
        {
            element.SendKeys(Keys.Tab);
        }
        //// ************************************VERTICAL_SCROLL**********
        public void WindowVerticalScroll(int pixel)
        {
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollBy(0, " + pixel + ")");
        }
        // ************************************Horizontal_SCROLL******************
        public void WindowHorizontalScroll()
        {
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollBy(2000, 0)");
        }
        //// scroll it at the top and then by a 1/4th of the height of the view
        public void WindowVerticalScroll(IWebElement element)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript(
                "arguments[0].scrollIntoView(true); window.scrollBy(0, -window.innerHeight / 4);", element);
        }
        //// scroll it at the bottom
        public void WindowScrollToBottom(IWebElement element)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(false);", element);
        }
        //// scroll to the element
        public void WindowScrollToTheElement(IWebElement element)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView();", element);

            try
            {
                Thread.Sleep(500);
            }
            catch (Exception e)
            {
                // Handle exception
                Console.WriteLine(e.ToString());
            }
        }
        // scroll just below an element , eg fixed header
        public void WindowScrollToBelowTheElement(IWebElement element, IWebElement header)
        {
            ((IJavaScriptExecutor)driver).ExecuteScript(
                "arguments[0].scrollIntoView(true); window.scrollBy(0, -arguments[1].offsetHeight);", element, header);
        }

        // *********************WAIT_TILL_ELEMENT_VISIBLE*********
        public void WaitTillElementVisible(IWebElement element)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(50));
            wait.Until(ExpectedConditions.ElementIsVisible(By.Id(element.GetAttribute("id"))));
        }
        // ******************************WAIT_SOME_TIME****************
        public void WaitForPageLoad(long milliseconds)
        {
            Thread.Sleep((int)milliseconds);
        }
        // *************************************ASSERT*************
        public bool Assert(string actualTitle, string expectedTitle)
        {
            return actualTitle.Equals(expectedTitle, StringComparison.OrdinalIgnoreCase);
        }
        // ******************************JAVASCRIPT_FUNCTIONS*****************************


        //Get TitleByJS
        public string GetTitleByJS(IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            return js.ExecuteScript("return document.title").ToString();
        }
        // to refresh webPage 
        public object RefreshWebPageByJS(IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            return js.ExecuteScript("history.go(0)");
        }
        // to send the value By JS
        public void SendKeysByJS(string id, string value, IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript($"document.getElementById('{id}').value = '{value}'");
        }

        // to click On element By Js
        public void ClickByJS(IWebElement element, IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].Click()", element);
        }
        // to scroll the bottom of the page
        public void ScrollDownPage(IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
        }

        // this method  is to perform  scrollUp page
        public void ScrollUpPage(IWebDriver driver, string high)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(0, '" + high + "')");
        }

        //flash
        public void Flash(IWebElement element, IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string bgColor = element.GetCssValue("background-color");

            for (int i = 0; i <= 10; i++)
            {
                ChangeColor("rgb(0, 200, 0)", element, driver);
                ChangeColor(bgColor, element, driver);
            }
        }

        private void ChangeColor(string color, IWebElement element, IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].style.backgroundColor = '" + color + "'", element);

            try
            {
                Thread.Sleep(100);
            }
            catch (ThreadInterruptedException e)
            {
                // Handle exception
            }
        }
        // to draw broder using js
        public void DrawBorder(IWebElement element)
        {
            IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
            jsExecutor.ExecuteScript("arguments[0].style.border='3px solid green'", element);
        }

        // alert message
        public void GenerateAlertMessage(IWebDriver driver, string message)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("alert('" + message + "')");
        }
        //	Js click on a Element
        public object ClickOnElementByJS(IWebDriver driver, IWebElement element)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            return js.ExecuteScript("arguments[0].click();", element);
        }

        //	to refresh the page using Js
        public void RefreshBrowserByJS(IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("history.go(0);");
        }
        //  to getTitle BY Js
        public string GetTitleByJs(IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string title = js.ExecuteScript("return document.title;").ToString();
            return title;
        }
        //this will return the page text
        public string GetPageInnerText(IWebDriver driver)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            string pageText = js.ExecuteScript("return document.documentElement.innerText;").ToString();
            return pageText;
        }

        //if data is get successfully then return row and column
        public string GetCurrentDate(string format)
        {
            // Create object of DateTimeFormatInfo class and decide the format
            DateTimeFormatInfo dateFormat = new DateTimeFormatInfo();

            // Get current date and time using DateTime.Now
            DateTime date = DateTime.Now;

            // Now format the date
            string date1 = date.ToString(format, dateFormat);

            return date1;
        }


    }









}











