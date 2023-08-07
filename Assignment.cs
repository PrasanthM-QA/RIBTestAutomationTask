using OpenQA.Selenium;
using POM.CommonMethods;
using SeleniumExtras.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POM.Source.Pages
{
    class Assignment
    {
        public IWebDriver driver;

        Base b;
       

        [FindsBy(How =How.XPath,Using = "(//li[@class='ui-widget-content ui-corner-tr ui-draggable ui-draggable-handle'])[1]")]
        public IWebElement source;

        [FindsBy(How = How.XPath, Using = "//div[@id='trash']")]
        public IWebElement target;

        [FindsBy(How = How.XPath, Using = "//li[@id='Photo Manager']")]
        public IWebElement photomanager;

        [FindsBy(How = How.XPath, Using = "//iframe[@class='demo-frame lazyloaded']")]
        public IWebElement frame;

        [FindsBy(How = How.XPath, Using = "//div[@class='menu-interaction-container']//span[.='Drag And Drop']")]
        public IWebElement drag;

        [FindsBy(How = How.XPath, Using = "//div[@class='menu-interaction-container']//span[.='DropDown Menu']")]
       public IWebElement dropdown;

        [FindsBy(How = How.XPath, Using = "//select")]
        public IWebElement dropdown1;

        [FindsBy(How = How.XPath, Using = "//div[@class='menu-miscellaneous-container']//span[.='Sample Page Test']")]
        public IWebElement samplePageTest;
        [FindsBy(How = How.XPath, Using = "//div[@id='dismiss-button']")]
        public IWebElement close;
        [FindsBy(How = How.XPath, Using = "//input[@name='g2599-name']")]
        public IWebElement name;

        [FindsBy(How = How.XPath, Using = "(//span[.='Progress Bar'])[2]")]
        public IWebElement progressbar;
        [FindsBy(How = How.XPath, Using = "//iframe[@id='ad_iframe']")]
        public IWebElement advertise;



        public Assignment(IWebDriver driver)
        {
            PageFactory.InitElements(driver, this);
            this.driver = driver;
            b = new Base(driver);
            driver.Manage().Timeouts().ImplicitWait=TimeSpan.FromSeconds(5);
    }
        
        public void dragAndDrop()
        {
            b.switchToElement(frame);
            b.highlightElement(source);
            b.highlightElement(target);
            b.dragAndDrop(source, target);
        }
        public void Dopdown()
        {
            b.switchToDefaultFrame();
            b.Click(dropdown);
            b.highlightElement(dropdown1);
            b.SelectByVisibleText(dropdown1, "Angola");
        }
        
        public void Tab()
        {
            b.switchToDefaultFrame();
            b.WindowScrollToTheElement(progressbar);
            b.highlightElement(samplePageTest);
            samplePageTest.Click();
            Thread.Sleep(30000);
            b.switchToElement(advertise);
            b.DismissAlert();
            b.highlightElement(close);
            b.Click(close);
           b.switchToDefaultFrame();
            Thread.Sleep(2000);
            b.highlightElement(name);
            name.SendKeys("ABC");
            b.tAB(name);
            
        }
    }
}
