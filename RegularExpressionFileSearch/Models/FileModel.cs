using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RegularExpressionFileSearch.Models
{
    public class FileModel
    {

        public string Text { get; set; }

        public List<string> RegexHits { get; set; }



        public string RegexPatteren { get; set; }

        public string FilePath { get; set; }


        public FileModel() 
        {
            RegexHits = new List<string>();
        }


        public override string ToString()
        {
            string info = "";

            foreach(string hit in RegexHits)
            {
                info += hit + "\n";
                   
            }

            info += "Text length: " + Text.Length + "\n" +
                    "Regex hits: " + RegexHits.Count + "\n";

            return info;
        }
    }
}
