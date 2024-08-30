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
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Bibliography;
using System.Runtime.CompilerServices;
using SixLabors.Fonts;
using System;
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

            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.marocannonces.com/categorie/309/Emploi/Offres-emploi.html");
            WebDriverWait wait = new(driver, TimeSpan.FromSeconds(10));
            IList<IWebElement> jobElements = driver.FindElements(By.CssSelector(".holder"));
            List<JobPosting> jobPostings = new List<JobPosting>();
            int i = 1;
            foreach (IWebElement element in jobElements)
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
            Console.WriteLine("^les informations de la page n°¨1 sont:");
            Afficher(jobPostings, i);



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


                // Populate the worksheet with data from the list of objects
                int row = 2;
                foreach (var jobPosting in jobPostings)
                {
                    worksheet.Cell(row, 1).Value = jobPosting.Title;
                    worksheet.Cell(row, 2).Value = jobPosting.Level;
                    worksheet.Cell(row, 3).Value = jobPosting.Salary;
                    worksheet.Cell(row, 4).Value = jobPosting.Location;

                    // Optional: Set default cell formatting for data rows
                    worksheet.Range(row, 1, row, 4).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Left); // Left align data

                    row++;
                }

                // les autres pages  
                 for (int j = 2; j <= 4; j++)
                 {
                     var link = "https://www.marocannonces.com/categorie/309/Emploi/Offres-emploi/@p.html";
                     driver.Navigate().GoToUrl(link.Replace("@p", j.ToString()));
                     // Attender que les elements se charge (auster le selected)
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
                     Console.WriteLine($"^les informations de la page n°¨{j} sont:");
                     Afficher(jobPostings, i = 1);

                    

                    // Populate the worksheet with data from the list of objects
                    
                    foreach (var jobPosting in jobPostings)
                    {
                        worksheet.Cell(row, 1).Value = jobPosting.Title;
                        worksheet.Cell(row, 2).Value = jobPosting.Level;
                        worksheet.Cell(row, 3).Value = jobPosting.Salary;
                        worksheet.Cell(row, 4).Value = jobPosting.Location;

                        // Optional: Set default cell formatting for data rows
                        worksheet.Range(row, 1, row, 4).Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)HorizontalAlignment.Left); // Left align data

                        row++;
                    }



                }

                // Save the workbook to a file
                workbook.SaveAs("job_postings.xlsx");
            }
            driver.Close();

        }
    }
}

        