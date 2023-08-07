using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Drawing.Imaging;

namespace POM.CommonMethods
{
    internal class Utility
    {
        public static long PAGE_LOAD_TIMEOUT = 20;
        public static long IMPLICIT_WAIT = 20;

        public static String TESTDATA_SHEET_PATH = "xlsx";

        public static IJavaScriptExecutor js;
        public static Actions act;
        public static IWebDriver driver;
        public String path;

        public Utility(IWebDriver driver)
        {
            Utility.driver = driver;
        }
        // *****************************READ_DATA_FROM_EXCEL************************************
        public object[][] ReadExcel(string filePath, string sheetName)
        {
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheet(sheetName);
                int rowCount = sheet.LastRowNum;
                int colCount = sheet.GetRow(0).LastCellNum;

                object[][] data = new object[rowCount][];

                for (int i = 0; i < rowCount; i++)
                {
                    data[i] = new object[colCount];
                    IRow row = sheet.GetRow(i + 1);
                    for (int j = 0; j < colCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        data[i][j] = cell.ToString();
                        Console.WriteLine(data[i][j]);
                    }
                }

                return data;
            }



        }
        // **********************************CAPTURE_SCREENSHOT_CURRENTPAGE*************************
        public string CaptureScreenshot(string filePath)
        {
            ITakesScreenshot screenshotDriver = (ITakesScreenshot)driver;
            Screenshot screenshot = screenshotDriver.GetScreenshot();

            string absolutePathLocation = Path.GetFullPath(filePath);
            screenshot.SaveAsFile(absolutePathLocation, ScreenshotImageFormat.Png);

            return absolutePathLocation;
        }
        // **********************************CAPTURE_SCREENSHOT_SELECTEDELEMENT********************
        public string CaptureScreenshotOfElement(IWebElement element, string filePath)
        {
            // Take a screenshot of the whole page
            Screenshot screenshot = driver.TakeScreenshot();

            // Get the location and size of the element
            int elementWidth = element.Size.Width;
            int elementHeight = element.Size.Height;
            Point elementLocation = element.Location;

            // Crop the screenshot to the element's dimensions
            using (MemoryStream ms = new MemoryStream(screenshot.AsByteArray))
            {
                System.Drawing.Image fullScreenImage = System.Drawing.Image.FromStream(ms);
                System.Drawing.Rectangle cropArea = new System.Drawing.Rectangle(elementLocation, new System.Drawing.Size(elementWidth, elementHeight));
                System.Drawing.Image croppedImage = new System.Drawing.Bitmap(cropArea.Width, cropArea.Height);

                using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(croppedImage))
                {
                    g.DrawImage(fullScreenImage, new System.Drawing.Rectangle(0, 0, cropArea.Width, cropArea.Height), cropArea, System.Drawing.GraphicsUnit.Pixel);
                }

                // Save the cropped image to the specified file path
                string absolutePathLocation = Path.GetFullPath(filePath);
                croppedImage.Save(absolutePathLocation, System.Drawing.Imaging.ImageFormat.Png);

                return absolutePathLocation;
            }

        }
        public string GetScreenshot(IWebDriver driver, string screenshotName)
        {
            string dateName = $"{DateTime.Now.Year:D4}{DateTime.Now.Month:D2}{DateTime.Now.Day:D2}{DateTime.Now.Hour:D2}{DateTime.Now.Minute:D2}{DateTime.Now.Second:D2}";
            ITakesScreenshot ts = (ITakesScreenshot)driver;
            Screenshot screenshot = ts.GetScreenshot();

            string directory = Path.Combine(Directory.GetCurrentDirectory(), "Screenshots");
            Directory.CreateDirectory(directory);

            string destination = Path.Combine(directory, screenshotName + dateName + ".png");

            screenshot.SaveAsFile(destination, ScreenshotImageFormat.Png);

            return destination;
        }
    }
}
