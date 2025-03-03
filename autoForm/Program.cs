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
            const int MINUTE_MIN = 1, MINUTE_MAX = 2;

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
            int rowIndex = 0;
            foreach (var row in rows)
            {
                rowIndex++;

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"===============Excel row number: {rowIndex}===============");
                Console.ForegroundColor = ConsoleColor.White;

                if (first)
                {
                    first = false;
                    continue;
                }

                bool success = false;
                int retryCount = 0;
                int innitAnsewer = ANSWER;

                while (!success && retryCount < 3) // Thử tối đa 3 lần
                {
                    try
                    {

                        // Navigate to the form page
                        driver.Navigate().GoToUrl("https://docs.google.com/forms/d/e/1FAIpQLSdtJj0aEPn4SdwQdNVq7-_Fs-qJAQXtd8ueSkA590v42jX24w/viewform");


                        #region page 1
                        // Diền Email
                        Console.WriteLine($"Page 1: NHAN DIEN DOI TƯƠNG");

                        IWebElement inputElement = driver.FindElement(By.CssSelector("input.whsOnd.zHQkBf"));
                        var ansInputText = ExcelHelpers.getAnswerString(row, ANSWER++);

                        // Gọi hàm nhập dữ liệu vào ô input
                        AnswerHelper.AnswerInputText(inputElement, ansInputText);

                        // end Điền Email
                        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        var spanElement = wait.Until(d => d.FindElement(By.XPath("//span[contains(text(), 'Có (Xin vui lòng tiếp tục trả lời phỏng vấn)')]")));

                        // Click on the span element
                        spanElement.Click();
                        Console.WriteLine($"Page 1: Choose Yes");

                        ANSWER++;

                        var wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                        var spanElement2 = wait.Until(d => d.FindElement(By.XPath("//span[contains(text(), 'Tiếp')]")));
                        spanElement2.Click();
                        Console.WriteLine($"Page 1: Next");

                        #endregion
                       

                        #region PAGE 2
                        #region 3 câu đơn đầu
                        Console.WriteLine($"Page 2: DETAILS");

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


                        var listNumAns = new List<int> { 5, 4, 4, 4, 4 };
                        var wait3 = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        var questionsMultil = wait3.Until(d => d.FindElements(By.CssSelector(".ufh7vf")));

                        //var questionsMultil = driver.FindElements(By.CssSelector(".ufh7vf"));
                        int i = 0;
                        foreach (var questions in questionsMultil)
                        {
                            var result = ExcelHelpers.getListAnswer(row, ANSWER, ANSWER + listNumAns[i] - 1);
                            AnswerHelper.AnswerMultipleQuestion(driver, questions, result);
                            ANSWER += listNumAns[i];
                            i++;
                        }
                        IWebElement nextButton = driver.FindElement(By.XPath("//span[contains(text(),'Tiếp')]"));
                        nextButton.Click();
                        IWebElement nextButton2 = driver.FindElement(By.XPath("//div[@role='button' and .//span[text()='Tiếp']]"));
                        nextButton2.Click();
                        Console.WriteLine($"Page 2: NEXT");
                        #endregion
                        #endregion

                        #region Page 3
                        Console.WriteLine($"Page 3: USER INFOMATION ");

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


                        success = true;

                        ANSWER = SKIP + 1;
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"SUCCESS");
                        Console.ForegroundColor = ConsoleColor.White;

                        int delay = rnd.Next(MINUTE_MIN , MINUTE_MAX ) * 1000;
                        Console.WriteLine($"⏳ Wait {delay / 1000 / 60} Minute {(delay / 1000) % 60}Second for next form...");
                        Thread.Sleep(delay);
                    }
                    catch (Exception ex)
                    {
                        retryCount++;
                        ANSWER = SKIP + 1;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Error row {rowIndex}: {ex.Message}");
                        Console.WriteLine($".............................");
                        Console.WriteLine($"Retry {retryCount} times");
                        Console.ForegroundColor = ConsoleColor.White;

                    }

                }

            }

            // Close the browser
            driver.Quit();
        }

    }
}
