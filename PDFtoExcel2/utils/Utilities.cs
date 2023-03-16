using System;
using System.Collections.Generic;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Filter;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Geom;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using Rectangle = iText.Kernel.Geom.Rectangle;
using Table = iText.Layout.Element.Table;
using Border = iText.Layout.Borders.Border;
using Cell = iText.Layout.Element.Cell;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Text;
using System.IO;
using System.Web.UI.WebControls;

namespace PDFtoExcel2.utils
{
    public class Utilities
    {

        private readonly List<string> dataset = new List<string>();

        public void PDFtoText(string fileName, PdfDocument document)
        {
            //assign PDF location to a string and create new StringBuilder...
            var pageText = new StringBuilder();

            //remove headers and footers
            Rectangle rect = new Rectangle(80, 0, 460, 800);
            TextRegionEventFilter regionFilter = new TextRegionEventFilter(rect);
            {
                var pageNumbers = document.GetNumberOfPages();
                for (int page = 1; page <= pageNumbers; page++)
                {
                    //create new text extraction renderer
                    //LocationTextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                    ITextExtractionStrategy strategy = new FilteredTextEventListener(new LocationTextExtractionStrategy(), regionFilter);
                    PdfCanvasProcessor parser = new PdfCanvasProcessor(strategy);
                    parser.ProcessPageContent(document.GetPage(page));
                    pageText.Append(strategy.GetResultantText() + "\n");
                }
                string extractedText = pageText.ToString().Trim();
                //Console.WriteLine(extractedText);

                //remove unwanted new lines
                extractedText = Regex.Replace(extractedText, @"((\b)\n(\s))", " ");

                //adding extracted text to dataset for conversion
                List<string> splittedText = extractedText.Split('\n').ToList();
                foreach (string x in splittedText)
                {
                    //remove other unnecessary rows
                    string pattern = @"^(Grand Total|Subgroup Total|Sub Group|Group|Trn|Code|\d*[0-9{4}](\s)\d)";
                    if (!Regex.IsMatch(x, pattern))
                    {
                        dataset.Add(x);
                    }
                }
            }
            convertToExcel(fileName, new XLWorkbook());
            convertToPDF(fileName);
            dataset.Clear();
        }

        public void convertToExcel(string fileName, XLWorkbook workbook)
        {
            StringBuilder pageText = new StringBuilder();

            for (int i = 0; i < dataset.Count; i++)
            {
                if (dataset[i].Length < 5)
                {
                    pageText.Append(dataset[i] + " " + dataset[i + 1] + "\n");
                    i++;
                }
                else if (i == dataset.Count - 1)
                {
                    pageText.Append(dataset[i]);
                }
                else
                {
                    pageText.Append(dataset[i] + "\n");
                }
            }

            char[] charSeparators = new char[] { '\n' };
            List<string> splittedData = pageText.ToString().Split(charSeparators).ToList();
            {
                IXLWorksheet worksheet = workbook.AddWorksheet("Sheet1");

                for (int i = 0; i < splittedData.Count; i++)
                {
                    var rowData = new List<string>();
                    StringBuilder sbID = new StringBuilder();

                    //filter ID
                    MatchCollection matchID = Regex.Matches(splittedData[i], @"^\d+[A-Za-z]?");
                    foreach (Match match in matchID)
                    {
                        sbID.Append(match.Value);
                    }

                    //filter decimals
                    MatchCollection matchDecimals = Regex.Matches(splittedData[i], @"(-\s)?\d{1,3}(?:,\d{3})*\.\d+");
                    StringBuilder sbDecimals = new StringBuilder();
                    foreach (Match match in matchDecimals)
                    {
                        sbDecimals.Append(match.Value + " ");
                    }
                    string decimals = sbDecimals.ToString().Trim();

                    //filter description
                    string description = splittedData[i].Replace(sbID.ToString(), "");
                    description = description.Replace(decimals, "");

                    //REMOVE WHITE SPACE BETWEEN NEGATIVE SIGN AND A DIGIT
                    string pattern = @"(?<=-)\s(?=\d)";
                    string rdecimals = Regex.Replace(decimals, pattern, "");

                    //adding data rows
                    rowData.Add(sbID.ToString());
                    rowData.Add(description.Trim());
                    List<string> sdecimals = rdecimals.Split().ToList();
                    foreach (var item in sdecimals)
                    {
                        rowData.Add(item);
                    }

                    //adding data rows in a cell
                    for (int j = 0; j < rowData.Count; j++)
                    {
                        worksheet.Cell(i + 1, j + 1).Value = rowData[j];
                    }
                }

                //make columns adjust to contents
                for (int i = 0; i < worksheet.Columns().Count(); i++)
                {
                    worksheet.Column(i + 1).AdjustToContents();
                }

                //save excel file to downloads folder
                //string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
                string filePath = System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\media\downloads\excel";
                char[] sp = { '.' };
                string[] newFileName = fileName.Split(sp);
                string excel_file_path = $@"{filePath}\{newFileName[0]}.xlsx";
                workbook.SaveAs(excel_file_path);
            }
        }

        public void convertToPDF(string fileName)
        {
            string filePath = System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\media\downloads\excel";
            char[] sp = { '.' };
            string[] newFileName = fileName.Split(sp);
            string excel_file_path = filePath + $"\\{newFileName[0]}.xlsx";

            var workbook = new XLWorkbook(excel_file_path);
            IXLWorksheet worksheet = workbook.Worksheet(1);

            string downloadsPath = System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\media\downloads\pdf";
            FileInfo file = new FileInfo($@"{downloadsPath}\{newFileName[0]}.pdf");
            file.Directory.Create();

            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(file));
            pdfDoc.SetDefaultPageSize(PageSize.A4.Rotate());
            Document doc = new Document(pdfDoc);

            Table table = new Table(UnitValue.CreatePercentArray(worksheet.Columns().Count())).UseAllAvailableWidth();
            table.SetBorder(Border.NO_BORDER);

            for (int i = 0; i < worksheet.Rows().Count(); i++)
            {
                for (int j = 0; j < worksheet.Columns().Count(); j++)
                {
                    Cell cell = new Cell().Add(new Paragraph(worksheet.Cell(i + 1, j + 1).GetValue<string>()));
                    table.AddCell(cell).SetFontSize(8);
                }
            }

            doc.Add(table);
            doc.Close();
        }
    }
}