using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using SixLabors.Fonts;
using DocumentFormat.OpenXml.Spreadsheet;
using WebDriverManager;


namespace consoleApp1
{

    public class JobPosting
    {
        public string Title { get; set; }
        public string Level { get; set; }
        public string Salary { get; set; }
        public string Location { get; set; }

        public JobPosting(IWebElement element)
        {
            // Implement logic to extract details from the element using appropriate selectors
            Title = element.FindElement(By.TagName("h3")).Text;
            Level = element.FindElement(By.ClassName("niveauetude")).Text;
            Salary = element.FindElement(By.ClassName("salary")).Text;
            Location = element.FindElement(By.ClassName("location")).Text;
        }
        private static void Afficher(List<JobPosting> jobPostings, int i)
        {


            foreach (var jobPosting in jobPostings)
            {

                Console.WriteLine($"Titre d'offre n{i}: {jobPosting.Title}");
                Console.WriteLine($"Niveau: {jobPosting.Level}");
                Console.WriteLine($"Salaire: {jobPosting.Salary}");
                Console.WriteLine($"Location: {jobPosting.Location}");
                Console.WriteLine();
                i++;
            }
        }
        static void Main(string[] args)
        {
            var opptions = new ChromeOptions();
            opptions.AddArgument("--headless");
            opptions.AddArgument("--disable-gpu");
            opptions.AddArgument("--window-size=1920,1080");
            opptions.AddArgument("--no-sandbox"); // Disable sandbox for headless mode


            // Initialize WebDriver
            using (IWebDriver driver = new ChromeDriver(opptions))
            {

                const int maxPages = 5;  // Adjust as needed
                int i = 1, row = 2;
                const string baseUrl = "https://www.marocannonces.com/categorie/309/Emploi/Offres-emploi.html";
                // Create a new workbook
                using (var workbook = new XLWorkbook())
                {
                    // Create a worksheet
                    var worksheet = workbook.Worksheets.Add("Job Postings");

                    // Add headers to the first row with formatting
                    worksheet.Cell(1, 1).Value = "Title";
                    worksheet.Cell(1, 1).Style.Font.SetBold(true); // Make title bold
                    worksheet.Cell(1, 1).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Center); // Center align Title
                    worksheet.Cell(1, 2).Value = "Level";
                    worksheet.Cell(1, 2).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Center); // Center align level
                    worksheet.Cell(1, 2).Style.Font.SetBold(true);
                    worksheet.Cell(1, 3).Value = "Salary";
                    worksheet.Cell(1, 3).Style.Font.SetBold(true);
                    worksheet.Cell(1, 3).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Center); // Center align salary
                    worksheet.Cell(1, 4).Value = "Location";
                    worksheet.Cell(1, 4).Style.Font.SetBold(true);
                    worksheet.Cell(1, 4).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Center); // Center align location

                    for (int page = 1; page <= maxPages; page++)
                    {
                        var jobPostings = new List<JobPosting>();
                        var url = " ";
                        if (page == 1)
                            url = baseUrl;
                        else
                        {
                            url = "https://www.marocannonces.com/categorie/309/Emploi/Offres-emploi/@p.html";
                            url = url.Replace("@p", page.ToString());
                        }

                        driver.Navigate().GoToUrl(url);
                        // Attender que les elements se charge (auster le selected)
                        WebDriverWait wait = new(driver, TimeSpan.FromSeconds(100));

                        IList<IWebElement> jobs = driver.FindElements(By.CssSelector(".holder"));

                        foreach (IWebElement element in jobs)
                        {
                            try
                            {
                                jobPostings.Add(new JobPosting(element));

                            }
                            catch (NoSuchElementException)
                            {
                                continue;
                            }
                        }
                        Console.WriteLine($"^les informations de la page n°¨{page} sont:");
                        Afficher(jobPostings, i);



                        // Populate the worksheet with data from the list of objects

                        foreach (var jobPosting in jobPostings)
                        {
                            worksheet.Cell(row, 1).Value = jobPosting.Title;
                            worksheet.Cell(row, 2).Value = jobPosting.Level;
                            worksheet.Cell(row, 3).Value = jobPosting.Salary;
                            worksheet.Cell(row, 4).Value = jobPosting.Location;
                            // Increase the width of columns 1 to 4
                            worksheet.Column(1).Width = 70; 
                            worksheet.Column(2).Width = 50;
                            worksheet.Column(3).Width = 40;
                            worksheet.Column(4).Width = 30;


                            // Set default cell formatting for data rows
                            worksheet.Range(row, 1, row, 4).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Left); // Left align data

                            row++;
                        }



                    }

                    // Save the workbook to a file
                    workbook.SaveAs("job_postings.xlsx");
                    Console.WriteLine("les informations sont bien enregistrees");
                }
                driver.Close();

            }
        }
    }
}


