using AventStack.ExtentReports;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using POM.Source.Pages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager;
using AventStack.ExtentReports.Reporter;
using Docker.DotNet.Models;
using NUnit.Framework.Interfaces;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Engine.ClientProtocol;
using RestSharp;
using RazorEngine.Compilation.ImpromptuInterface.InvokeExt;

namespace POM.Test.HomePageTest
{
    
       class HomePageTest : IDisposable
    {
        public Assignment hp;
        public IWebDriver driver;
       
        
        [OneTimeSetUp]
        public void BrowserSetup()
            
        {
            new DriverManager().SetUpDriver(new ChromeConfig());
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "https://www.globalsqa.com/demo-site/draganddrop/";

        }


        /*
        [Test , Order(1)]
       public void DragAndDrop()
        {

            hp = new Assignment(driver);
            hp.dragAndDrop();
        }
        
        [Test ,Order(2)]
        public void dropdown()
        {

            hp = new Assignment(driver);
            hp.Dopdown();
        }
        */
        [Test, Order(3)]
        public void UseTab()
        {

            hp = new Assignment(driver);
            hp.Tab();
        }


        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}



