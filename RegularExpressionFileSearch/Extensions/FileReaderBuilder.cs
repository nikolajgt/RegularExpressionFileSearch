using IronOcr;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using RegularExpressionFileSearch.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RegularExpressionFileSearch.Extensions
{
    public class FileReaderBuilder
    {
        private FileModel _file = new FileModel();


        public FileReaderBuilder SetPath(string path)
        {
            _file.FilePath = path;
            return this;
        }

        public FileReaderBuilder SetRegexPattern(string pattern)
        {
            _file.RegexPatteren = pattern;
            return this;
        }




        public FileReaderBuilder GetWordContent()
        {
            string text = "";

            try
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

                Document doc = new Document();

                doc = app.Documents.Open(_file.FilePath);

                text = doc.Content.Text;

                doc.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n" + ex.Message);
            }

            _file.Text = text;

            return this;
        }


        public FileReaderBuilder GetPdfContent()
        {
            string content = "";

            try
            {
                PdfDocument pdf = new PdfDocument(new PdfReader(_file.FilePath));

                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var page = pdf.GetPage(i);

                    content += PdfTextExtractor.GetTextFromPage(page);

                }

                pdf.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n" + ex.Message);
            }

            _file.Text = content.ToString();

            return this;
        }



        public FileReaderBuilder GetImageTextContent(OcrLanguage ocrLanguage)
        {
            string text = "";

            try
            {
                IronTesseract ocr = new IronTesseract();
                ocr.Language = ocrLanguage;

                using (var Input = new OcrInput(_file.FilePath))
                {

                    // Input.Deskew();  // use if image not straight
                    // Input.DeNoise(); // use if image contains digital noise

                    OcrResult Result = ocr.Read(Input);

                    text = Result.Text;

                    text = Regex.Replace(text, @"\s+", "");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n" + ex.Message);
            }

            _file.Text = text;

            return this;
        }




        public FileReaderBuilder GetExcelContent()
        {
            string excelData = "";

            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                Workbook workbook = app.Workbooks.Open(_file.FilePath);
                Worksheet worksheet = workbook.Worksheets[1];

                Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        excelData += ((range.Cells[row, col] as Microsoft.Office.Interop.Excel.Range).Value2 ?? "").ToString() + "\t";
                    }
                    excelData += "\n";
                }

                workbook.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n" + ex.Message);
            }

            _file.Text = excelData;

            return this;
        }




        public FileReaderBuilder GetContentFromEnumType(FileTypeEnum type)
        {
            FileReaderBuilder placeholder = new FileReaderBuilder();

            switch (type)
            {
                case FileTypeEnum.Word:
                    return GetWordContent();


                case FileTypeEnum.Excel:
                    return GetExcelContent();


                case FileTypeEnum.Pdf:
                    return GetPdfContent();


                case FileTypeEnum.Image:
                    return GetImageTextContent(OcrLanguage.English);

            }
            return placeholder;
        }




        public FileReaderBuilder FindRegexHits()
        {
            try
            {
                MatchCollection match = Regex.Matches(_file.Text, _file.RegexPatteren);

                if (match.Count > 0)
                {
                    foreach (Match m in match)
                    {
                        _file.RegexHits.Add(m.Value);
                    }
                }
                else
                {
                    Console.WriteLine("No match found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return this;
        }

        public override string ToString()
        {
            return _file.ToString();
        }
    }
}
