using System;
using System.Threading.Tasks;
using autoForm.helpers;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace autoForm
{
    class Program
    {
        static async Task Main(string[] args)
        {
            #region Khởi tạo

            const int NUM_SIGLE_QUE = 3;
            const int NUM_MULTY_QUE = 3;

            int MAX_ANSWER = 31;
            int SKIP = 0;
            int ANSWER = SKIP + 1;

            Random rnd = new Random();
            #endregion

            // Load the Excel file
            using var workbook = new XLWorkbook("C:\\Users\\vuyou\\OneDrive\\Desktop\\DEVELOP\\autoForm\\autoForm\\data.xlsx");
            var worksheet = workbook.Worksheet(1);
            var rows = worksheet.RangeUsed().RowsUsed();

            // Initialize Selenium WebDriversdd4
            using var driver = new ChromeDriver();
            bool first = true;
            foreach (var row in rows)
            {
                if (first)
                {
                    first = false;
                    continue;
                }
                // Navigate to the form page
                driver.Navigate().GoToUrl("https://docs.google.com/forms/d/e/1FAIpQLSdtJj0aEPn4SdwQdNVq7-_Fs-qJAQXtd8ueSkA590v42jX24w/viewform");


                #region Nhận diện đối tượng
                // Diền Email

                IWebElement inputElement = driver.FindElement(By.CssSelector("input.whsOnd.zHQkBf"));
                var ansInputText = ExcelHelpers.getAnswerString(row, ANSWER++);

                // Gọi hàm nhập dữ liệu vào ô input
                AnswerHelper.AnswerInputText(inputElement, ansInputText);

                // end Điền Email
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                var spanElement = wait.Until(d => d.FindElement(By.XPath("//span[contains(text(), 'Có (Xin vui lòng tiếp tục trả lời phỏng vấn)')]")));

                // Click on the span element
                spanElement.Click();
                ANSWER++;

                var wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                var spanElement2 = wait.Until(d => d.FindElement(By.XPath("//span[contains(text(), 'Tiếp')]")));
                spanElement2.Click();

                #endregion
                // Iterate through the rows in the Excel file and fill the form

                // 3 câu đơn
                #region 3 câu đơn đầu
                var questionsSigles = driver.FindElements(By.CssSelector(".SG0AAe"));
                foreach (var question in questionsSigles)
                {
                    var result = ExcelHelpers.getAnswer(row, ANSWER);
                    AnswerHelper.AnswerSingleQuestion(question, result);
                    ANSWER++;
                }
                #endregion

                #region 5 caua nhieeuf
                IJavaScriptExecutor jsPage2 = (IJavaScriptExecutor)driver;
                wait.Until(d => jsPage2.ExecuteScript("return document.readyState").ToString() == "complete");


                var listNumAns = new List<int> {5 ,4, 4, 4, 4};
                var wait3 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                var questionsMultil = wait3.Until(d => d.FindElements(By.CssSelector(".ufh7vf")));

                //var questionsMultil = driver.FindElements(By.CssSelector(".ufh7vf"));
                int i = 0;
                foreach (var questions in questionsMultil)
                {
                    var result = ExcelHelpers.getListAnswer(row, ANSWER, ANSWER + listNumAns[i] - 1);
                    AnswerHelper.AnswerMultipleQuestion(driver,questions, result);
                    ANSWER += listNumAns[i] ;
                    i++;
                }

                #endregion

                var s = 2;


                IWebElement nextButton = driver.FindElement(By.XPath("//span[contains(text(),'Tiếp')]"));
                nextButton.Click();
                IWebElement nextButton2 = driver.FindElement(By.XPath("//div[@role='button' and .//span[text()='Tiếp']]"));
                nextButton2.Click();

                #region Page 3
                IJavaScriptExecutor jsPage3 = (IJavaScriptExecutor)driver;
                wait.Until(d => jsPage3.ExecuteScript("return document.readyState").ToString() == "complete");

                #region 3 câu đơn đầu
                var wait4 = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                var page3QuestionsSigles = wait4.Until(d => d.FindElements(By.CssSelector(".SG0AAe")));

                foreach (var question in page3QuestionsSigles)
                {
                    var result = ExcelHelpers.getAnswer(row, ANSWER);
                    AnswerHelper.AnswerSingleQuestion(question, result);
                    ANSWER++;
                }


                #endregion
                IWebElement submitButton = driver.FindElement(By.XPath("//div[@role='button' and .//span[text()='Gửi']]"));
                submitButton.Click();

                #endregion
                ANSWER = SKIP+1;

                int delay = rnd.Next(10, 60) * 1000;
                Console.WriteLine($"⏳ Wait {delay / 1000} Second for next form...");
                Thread.Sleep(delay);
            }

            // Close the browser
            driver.Quit();
        }


        // convert number to selecttion

        // funtion get sigle quétion

        // funtion get multi quétion
    }
}
