using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using IronOcr;
using Microsoft.Office.Interop.Word;
using RegularExpressionFileSearch.Extensions;
using RegularExpressionFileSearch.Models;


namespace RegularExpressionOffice
{
    class Program
    {

        /// Den valgte regex pattern der skal bruges
        private static readonly string regexPattern = @"(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])\d{2}([-]?| - | -| -)\d{4}";


        /// Her kan du vælge hvad du gerne vil scanne
        private static readonly FileTypeEnum fileType = FileTypeEnum.Pdf;


        /// En nem måde at få path ud via en enum
        private static readonly Dictionary<FileTypeEnum, string> fileCollection = new Dictionary<FileTypeEnum, string>
        {
            {FileTypeEnum.Word, @"D:\dummy\cpr.docx"},
            {FileTypeEnum.Excel, @"D:\dummy\cpr.xlsx"},
            {FileTypeEnum.Pdf, @"D:\dummy\cpr2.pdf"},
            {FileTypeEnum.Image, @"D:\dummy\cpr.png"}
        };


        /// Brug den her hvis du vælger GetImageContent(english). GetContentFromEnumType vælger english ved default.
        private static readonly OcrLanguage english = OcrLanguage.English;


        /// Har lavet det i form af builder pattern
        static void Main(string[] args)
        {

            string path = fileCollection.FirstOrDefault(x => x.Key == fileType).Value;


            FileReaderBuilder builder = new FileReaderBuilder()
                                              .SetPath(path)                                            /// Sætter path til den fil der skal undersøges
                                              .SetRegexPattern(regexPattern)                            /// Her kan du sætte Regex Patteren den skal findes
                                              .GetContentFromEnumType(fileType)                         /// Den her kan skiftes ud med noget mere specifikt f.eks GetImageContent(OcrLanguage lang) eller bare GetWordContent()
                                              .FindRegexHits();                                         /// Den her skal man sætte efter GetContent for at finde det man vil lede efter.



            Console.WriteLine(builder.ToString());

        }























        public void create_cpr_number()
        {
            Random rand = new Random();
            List<string> cpr = new List<string>();

            for (int i = 0; i < 999; i++)
            {
                var day = rand.Next(1, 31);
                var month = rand.Next(1, 12);
            }
        }

        public static void test()
        {
            List<string> list = new List<string>();
            Random rand = new Random();

            for (int i = 0; i < 999; i++)
            {
                var test = DateTime.Now.ToString("dd/mm/yy");

                test = test.Replace(@"-", String.Empty);

                var day = rand.Next(1000, 9999);

                Console.WriteLine(test + "-" + day);
                list.Add(test + "-" + day);
            }


        }

    }
}