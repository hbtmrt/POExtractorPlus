using SautinSoft;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POExtractorPlus.ExtractionBehavior
{
    public class Common
    {
        internal static Dictionary<string, string> GetFileContents(string[] files)
        {
            Dictionary<string, string> contents = new Dictionary<string, string>();

            foreach (var file in files)
            {
                string content = string.Empty;

                PdfFocus f = new PdfFocus();
                f.XmlOptions.ConvertNonTabularDataToSpreadsheet = true;

                f.OpenPdf(file);

                if (f.PageCount > 0)
                {
                    content = f.ToXml();
                }

                contents.Add(file, content);
            }

            return contents;
        }

        public List<string> PossibleSizes { get; set; }

        public Common() {
            SetPossibleSizes();
        }

        private void SetPossibleSizes()
        {
            this.PossibleSizes = new List<string>();

            this.PossibleSizes.Add("");
        }
    }
}
