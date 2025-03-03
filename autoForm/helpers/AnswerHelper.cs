
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;

namespace autoForm.helpers
{
    public static class AnswerHelper
    {
        public static void AnswerSingleQuestion(IWebElement question, int answer)
        {
            try
            {
                // Lấy tất cả các tùy chọn trong câu hỏi này
                var options = question.FindElements(By.CssSelector(".docssharedWizToggleLabeledContainer"));

                // Kiểm tra nếu có ít nhất 4 phần tử
                if (options.Count > answer)
                {
                    // Click vào phần tử thứ 4 (index 3)
                    options[answer-1].Click();
                    Console.WriteLine($"✅ Choose ansewer {answer}");
                }
                else
                {
                    options[options.Count -1].Click();
                }
            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine($"Lỗi: {ex.Message}");
            }

        }

        public static void AnswerMultipleQuestion(ChromeDriver driver, IWebElement question, List<int> answer)
        {
            try
            {
              //  IWebElement questionElement = driver.FindElement(By.CssSelector(".ufh7vf")); // Thay selector chính xác

                var options = question.FindElements(By.XPath(".//div[@role='radio' or @role='checkbox']"));

                if (options.Count == 0)
                {
                    throw new Exception("Không tìm thấy lựa chọn nào trong câu hỏi!");
                }

                // Lặp qua danh sách chỉ mục cần chọn
                var interval = 0;
                foreach (int index in answer)
                {
                    if (index < 0 || index >= options.Count)
                    {
                        Console.WriteLine($"⚠️ Cảnh báo: Index {index} không hợp lệ.");
                        continue;  // Bỏ qua nếu index không hợp lệ
                    }

                    IWebElement option = options[index + interval-1];

                    // Kiểm tra xem lựa chọn có bị disabled không
                    string isDisabled = option.GetAttribute("aria-disabled");
                    if (isDisabled == "true")
                    {
                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                        js.ExecuteScript("arguments[0].removeAttribute('aria-disabled');", option);

                    }

                    // Click vào lựa chọn
                    Actions actions = new Actions(driver);
                    actions.MoveToElement(option).Click().Perform();
                    Console.WriteLine($"✅ Choose ansewer {index + 1}");
                    interval += 5;
                }

            }
            catch (NoSuchElementException ex)
            {
                Console.WriteLine($"Lỗi: {ex.Message}");
            }

        }

        public static void AnswerInputText(IWebElement question, string answer)
        {
            if (question != null && question.Displayed && question.Enabled)
            {
                question.Clear();  // Xóa dữ liệu cũ (nếu có)
                question.SendKeys(answer);  // Nhập dữ liệu mới
            }
            else
            {
                throw new Exception("Element is either null, not visible, or disabled.");
            }
        }
    }
}
