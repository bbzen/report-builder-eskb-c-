using ClosedXML.Excel;
using HtmlAgilityPack;
using System.Text;

namespace MyService
{
    public class FileConverter
    {
        public void ConverToXls()
        {
            string[] foundFiles = FindFiles();
            foreach (string file in foundFiles)
            {
                convertFileToXlsBook(file);
            }
        }

        private async void convertFileToXlsBook(string filePath)
        {
            //read existing .html file
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            HtmlDocument doc = new HtmlDocument();
            //set the right encoding for loading html doc
            doc.OptionDefaultStreamEncoding = Encoding.GetEncoding("windows-1251");
            doc.Load(filePath);

            //prepare a new Excel book
            XLWorkbook currentBook = new XLWorkbook();
            IXLWorksheet currentWorkSheet = currentBook.Worksheets.Add("ER-Telecom");


            //add headers into the sheet
            HtmlNodeCollection bodyTableHeader = doc.DocumentNode.SelectNodes("//body/table/thead/tr/th");

            for (int i = 0; i < bodyTableHeader.Count; i++)
            {
                currentWorkSheet.Cell(2, i + 1).Value = bodyTableHeader[i].InnerText;
            }

            //add data from main table into sheet
            HtmlNodeCollection bodyDataRow = doc.DocumentNode.SelectNodes("//body/table/tbody/tr/td");
            int htmlNodeCount = bodyDataRow.Count / 10;

            for (int i = 0; i < htmlNodeCount; i++)
            {
                int row = i + 3;
                for (int j = 0; j < 10; j++)
                {
                    int htmlDataCell = i * 10 + j;
                    if (j == 6)
                    {
                        currentWorkSheet.Cell(row, j + 1).Value = DateTime.Parse(bodyDataRow[htmlDataCell].InnerText);

                    }
                    else if (0 < j && j < 5)
                    {
                        currentWorkSheet.Cell(row, j + 1).Value = Double.Parse(bodyDataRow[htmlDataCell].InnerText.Replace('.', ','));

                    }
                    else if (j == 0 || j == 7 || j == 9)
                    {
                        currentWorkSheet.Cell(row, j + 1).Value = Int64.Parse(bodyDataRow[htmlDataCell].InnerText);
                    }
                    else
                    {
                        currentWorkSheet.Cell(row, j + 1).Value = bodyDataRow[htmlDataCell].InnerText;
                    }
                }
            }
            //autoadjust the width of cell to data
            currentWorkSheet.Columns().AdjustToContents();

            //add title into the sheet
            HtmlNodeCollection bodyTableTitleFirst = doc.DocumentNode.SelectNodes("//body/h1");
            HtmlNodeCollection bodyTableTitleSecond = doc.DocumentNode.SelectNodes("//body/h2");
            string resultTitle = "";
            foreach (var item in bodyTableTitleFirst)
            {
                resultTitle += item.InnerText;
            }
            resultTitle += " Узел учета: ";
            foreach (var item in bodyTableTitleSecond)
            {
                resultTitle += item.InnerText;
            }
            currentWorkSheet.Cell(1, 1).Value = resultTitle;

            //saving the book
            try
            {
                currentBook.SaveAs(filePath + ".xlsx");

            }
            catch (IOException e)
            {
                Console.WriteLine();
                Console.WriteLine("Ошибка! Открыт файл с готовым отчетом. Закройте файл отчета и перезапустите программу.");
            }
        }

        private string[] FindFiles()
        {
            string[] foundFiles = Directory.GetFiles("./reports/", "*.html");

            if (foundFiles != null)
            {
                Console.WriteLine("Found files: ");
                foreach (var item in foundFiles)
                {
                    Console.WriteLine(item);
                }
            }
            else
            {
                throw new FileNotFoundException("No apropriate files in \"./reports/\" has been found");
            }
            return foundFiles;
        }
    }
}