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

partial class Program
{
    static void Main(string[] args)
    {

        IWebDriver driver = new ChromeDriver();
        driver.Navigate().GoToUrl("https://www.marocannonces.com/categorie/309/Emploi/Offres-emploi.html");
        WebDriverWait wait = new(driver, TimeSpan.FromSeconds(10));
        IList<IWebElement> jobElements = driver.FindElements(By.CssSelector(".holder"));
        int j = 0;


        // Créer un fichier Excel
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Offres d'emploi");

            // Ajouter les en-têtes
            worksheet.Cell(1, 1).Value = "Titre de l'offre";
            worksheet.Cell(1, 2).Value = "Niveau d'étude";
            worksheet.Cell(1, 3).Value = "Salaire";
            worksheet.Cell(1, 4).Value = "Location";
            int row = 2;

            foreach (IWebElement otheroffre in jobElements)

            {
                j++;
                IWebElement titreOffre = otheroffre.FindElement(By.TagName("h3"));

                IWebElement niveauetude = otheroffre.FindElement(By.ClassName("niveauetude"));
                IWebElement salaire = otheroffre.FindElement(By.ClassName("salary"));
                IWebElement location = otheroffre.FindElement(By.ClassName("location"));
                Console.WriteLine($"les informations de la page N°{1}sont:");
                Console.WriteLine($"Titre d'offre N°{j}: " + titreOffre.Text);
                Console.WriteLine("Niveau d'étude: " + niveauetude.Text);
                Console.WriteLine("Salaire: " + salaire.Text);
                Console.WriteLine("Location: " + location.Text);

                // Stocker les informations dans le fichier Excel
                worksheet.Cell(row, 1).Value = titreOffre.Text;
                worksheet.Cell(row, 2).Value = niveauetude.Text;
                worksheet.Cell(row, 3).Value = salaire.Text;
                worksheet.Cell(row, 4).Value = location.Text;

                row++;

            }
            TimeSpan.FromSeconds(1000);

            for (int i = 2; i <= 4; i++)
            {
                var link = "https://www.marocannonces.com/categorie/309/Emploi/Offres-emploi/@p.html";
                driver.Navigate().GoToUrl(link.Replace("@p", i.ToString()));

                // Attender que les elements se charge (auster le selected)
                IList<IWebElement> jobs = driver.FindElements(By.CssSelector(".holder"));
                row++;
                foreach (IWebElement offre in jobs)

                {

                    IWebElement titreOffre = offre.FindElement(By.TagName("h3"));

                    IWebElement niveauetude = offre.FindElement(By.ClassName("niveauetude"));
                    IWebElement salaire = offre.FindElement(By.ClassName("salary"));
                    IWebElement location = offre.FindElement(By.ClassName("location"));
                    Console.WriteLine($"les informations de la page N°{i}sont:");
                    Console.WriteLine($"Titre d'offre N°{j}: " + titreOffre.Text);
                    Console.WriteLine("Niveau d'étude: " + niveauetude.Text);
                    Console.WriteLine("Salaire: " + salaire.Text);
                    Console.WriteLine("Location: " + location.Text);
                    // Stocker les informations dans le fichier Excel
                    worksheet.Cell(row, 1).Value = titreOffre.Text;
                    worksheet.Cell(row, 2).Value = niveauetude.Text;
                    worksheet.Cell(row, 3).Value = salaire.Text;
                    worksheet.Cell(row, 4).Value = location.Text;

                    row++;
                    j++;
                }

            }
            TimeSpan.FromSeconds(6000);
            workbook.SaveAs(" les Offres_emploi.xlsx");
        }
        driver.Close();

        Console.WriteLine("les donnees sont enregister ");


    }

}