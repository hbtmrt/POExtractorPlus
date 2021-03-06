﻿using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
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
        public enum ExtractionTechnology
        {
            Itext,
            Sautin
        }

        internal static Dictionary<string, string> GetFileContents(string[] files, ExtractionTechnology technology)
        {
            switch (technology) {
                case ExtractionTechnology.Itext:
                    return GetFileContentsThroughIText(files);
                case ExtractionTechnology.Sautin:
                    return GetFileContentsThroughSautin(files);
                default:
                    return GetFileContentsThroughIText(files);
            }
        }

        private static Dictionary<string, string> GetFileContentsThroughIText(string[] files) {
            Dictionary<string, string> contents = new Dictionary<string, string>();

            foreach (var file in files)
            {
                StringBuilder text = new StringBuilder();

                using (PdfReader reader = new PdfReader(file))
                {
                    ITextExtractionStrategy Strategy = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();

                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        string page = PdfTextExtractor.GetTextFromPage(reader, i, Strategy);
                        text.Append(page);
                    }

                }

                contents.Add(file, text.ToString());
            }

            return contents;
        }
        
        private static Dictionary<string, string> GetFileContentsThroughSautin(string[] files)
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

        public static List<string> PossibleSizes { get; set; }

        public Common() {
        }

        public static List<string> GetPossibleSizes()
        {
            var possibleSizes = new List<string>();

            possibleSizes.Add("XS");
            possibleSizes.Add("S");
            possibleSizes.Add("M");
            possibleSizes.Add("L");
            possibleSizes.Add("XL");
            possibleSizes.Add("XXL");
            possibleSizes.Add("ECH");
            possibleSizes.Add("CH");
            possibleSizes.Add("M");
            possibleSizes.Add("G");
            possibleSizes.Add("EG");
            possibleSizes.Add("EEG");
            possibleSizes.Add("85");
            possibleSizes.Add("90");
            possibleSizes.Add("95");
            possibleSizes.Add("100");
            possibleSizes.Add("105");
            possibleSizes.Add("110");
            possibleSizes.Add("80");
            possibleSizes.Add("85");

            return possibleSizes;
        }
    }
}
