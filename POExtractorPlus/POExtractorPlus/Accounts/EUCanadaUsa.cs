using ClosedXML.Excel;
using POExtractorPlus.ExtractionBehavior;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace POExtractorPlus.Accounts
{
    public class EUCanadaUsa : IAccount
    {
        public Dictionary<string, ExtractedRowModel> ExtractedLines { get; set; }
        public List<string> RejectedFiles { get; set; }
        Thread exportThread;
        public string Destination { get; set; }

        public EUCanadaUsa()
        {
            this.ExtractedLines = new Dictionary<string, ExtractedRowModel>();
            this.RejectedFiles = new List<string>();
            exportThread = new Thread(Export);
        }

        public List<string> Extract(string destination, string[] files)
        {
            this.Destination = destination;
            Dictionary<string, string> fileContents = Common.GetFileContents(files);

            foreach (var item in fileContents.ToList())
            {
                ExtractedRowModel extractedLine = ExtractFile(item.Key, item.Value);
                this.ExtractedLines.Add(item.Key, extractedLine);
            }

            exportThread.Start();

            return this.RejectedFiles;
        }

        public void Export()
        {
            var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("PO Details");

            ws.Cell("A1").Value = "PO Number";
            ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("A1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("A1:A2").Merge();

            ws.Cell("B1").Value = "Manufacturer";
            ws.Cell("B1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("B1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("B1:B2").Merge();

            ws.Cell("C1").Value = "Season code";
            ws.Cell("C1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("C1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("C1:C2").Merge();

            ws.Cell("D1").Value = "Material";
            ws.Cell("D1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("D1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("D1:D2").Merge();

            ws.Cell("E1").Value = "Material Description";
            ws.Cell("E1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("E1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("E1:E2").Merge();

            ws.Cell("F1").Value = "Plan Ex-fty";
            ws.Cell("F1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("F1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("F1:F2").Merge();

            ws.Cell("G1").Value = "Original Ex-fty";
            ws.Cell("G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("G1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("G1:G2").Merge();

            ws.Cell("H1").Value = "Total PO qty";
            ws.Cell("H1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("H1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("H1:H2").Merge();

            ws.Cell("I1").Value = "PO unit price";
            ws.Cell("I1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("I1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("I1:I2").Merge();

            ws.Cell("J1").Value = "Delivery Address";
            ws.Cell("J1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("J1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("J1:J2").Merge();

            ws.Cell("K1").Value = "Trans Mode";
            ws.Cell("K1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell("K1").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("K1:K2").Merge();

            ws.Cell("L1").Value = "Size";

            Dictionary<string, int> sizeMapping = new Dictionary<string, int>();
            int sizeKeyIndex = 76;

            var row = 3;

            foreach (var key in ExtractedLines.Keys)
            {
                var line = ExtractedLines[key];

                if (!line.HasErrors)
                {
                    string cellName = "A" + row;
                    ws.Cell(cellName).Value = line.PONumber;

                    cellName = "B" + row;
                    ws.Cell(cellName).Value = line.Manufacturer;

                    cellName = "C" + row;
                    ws.Cell(cellName).Value = line.SeasonCode;

                    cellName = "D" + row;
                    ws.Cell(cellName).Value = line.Material;

                    cellName = "E" + row;
                    ws.Cell(cellName).Value = line.MaterialDescription;

                    cellName = "F" + row;
                    ws.Cell(cellName).Value = line.Plan_ExFTY;

                    cellName = "G" + row;
                    ws.Cell(cellName).Value = line.Original_ExFTY;

                    cellName = "H" + row;
                    ws.Cell(cellName).Value = line.TotalPOQty;

                    cellName = "I" + row;
                    ws.Cell(cellName).Value = line.POUnitPrice;

                    cellName = "J" + row;
                    ws.Cell(cellName).Value = line.DeliveryAddress;

                    cellName = "K" + row;
                    ws.Cell(cellName).Value = line.TransMode;

                    foreach (var sizeKey in line.Sizes.Keys)
                    {
                        int sizeColumnNo = 0;
                        if (sizeMapping.ContainsKey(sizeKey))
                        {
                            sizeColumnNo = sizeMapping[sizeKey];
                        }
                        else
                        {
                            sizeMapping.Add(sizeKey, sizeKeyIndex);

                            string newSizeColumnName = Char.ConvertFromUtf32(sizeKeyIndex) + "2";
                            ws.Cell(newSizeColumnName).Value = sizeKey;
                            ws.Cell(newSizeColumnName).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            sizeColumnNo = sizeKeyIndex;
                            sizeKeyIndex++;
                        }

                        var sizeValue = line.Sizes[sizeKey];
                        string sizeValueColumnName = Char.ConvertFromUtf32(sizeColumnNo) + row;
                        ws.Cell(sizeValueColumnName).Value = sizeValue;
                    }

                    row++;
                }
            }

            string sizeRangeColumn = "L1:" + Char.ConvertFromUtf32(sizeKeyIndex - 1) + "1";

            ws.Cell("L1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(sizeRangeColumn).Merge();

            //Adjust column widths to their content
            ws.Columns().AdjustToContents();

            ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

            string headerRange = "A1:" + Char.ConvertFromUtf32(sizeKeyIndex - 1) + "2";
            var rngTable = ws.Range(headerRange);
            var rngHeaders = rngTable.Range(headerRange); // The address is relative to rngTable (NOT the worksheet)
            rngHeaders.Style.Font.Bold = true;

            string destination = string.Format(@"{0}\{1}", this.Destination, "PODetails_" + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_tt")) + ".xlsx"; ;
            workbook.SaveAs(destination);
        }

        private ExtractedRowModel ExtractFile(string fileName, string fileXmlContent)
        {
            ExtractedRowModel model = new ExtractedRowModel();

            XmlSerializer serializer = new XmlSerializer(typeof(PdfXmlMapper));
            PdfXmlMapper mapping = null;

            using (TextReader reader = new StringReader(fileXmlContent))
            {
                mapping = (PdfXmlMapper)serializer.Deserialize(reader);
            }

            if (mapping != null)
            {
                try
                {
                    // Extract PO Number
                    model.PONumber = ExtractPONumber(mapping);
                    model.SeasonCode = ExtractSeasonCode(mapping);
                    model.Manufacturer = ExtractManufacturer(mapping);
                    model.Material = ExtractMaterial(mapping);
                    model.MaterialDescription = ExtractMaterialDescription(mapping);
                    model.Plan_ExFTY = ExtractPlannedDelDate(mapping);
                    model.Original_ExFTY = ExtractOriginalExFacDate(mapping);
                    model.TotalPOQty = ExtractTotalPOLineQty(mapping);
                    model.POUnitPrice = ExtractPOUnitPrice(mapping);
                    model.DeliveryAddress = ExtractDeliveryAddress(mapping);
                    model.TransMode = ExtractTransMode(mapping);
                    model.Sizes = ExtractSizes(mapping);
                }
                catch (Exception ex)
                {
                    model.HasErrors = true;
                    model.Error = ex.Message;

                    this.RejectedFiles.Add(fileName);
                }
            }

            return model;
        }

        private string ExtractPONumber(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            var cell = row.Cells[i];
                            if (cell.Equals(Constants.Common.PONumber))
                            {
                                string poNumber = row.Cells[i + 1];
                                return poNumber;
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractSeasonCode(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            var cell = row.Cells[i];
                            if (cell.Equals(Constants.Common.SeasonCode))
                            {
                                string code = row.Cells[i + 1];
                                return code;
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractManufacturer(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    foreach (var row in table.Rows)
                    {
                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            var cell = row.Cells[i];
                            if (cell.Equals(Constants.Common.Manufacturer))
                            {
                                string manufacturer = row.Cells[i + 1];
                                return manufacturer;
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractMaterial(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.Material))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractMaterialDescription(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.MaterialDescription))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractDeliveryAddress(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.DeliveryAddress))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractPannedExFacDate(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.PannedExFacDate))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractOriginalExFacDate(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.OriginalExFacDate))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractPOUnitPrice(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.POUnitPrice))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractTotalPOLineQty(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.TotalPOLineQty))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractPlannedDelDate(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.PlannedDelDate))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private string ExtractTransMode(PdfXmlMapper mapping)
        {
            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];

                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            var cell = row.Cells[j];
                            if (cell.Equals(Constants.Common.TransMode))
                            {
                                var nextRow = table.Rows[i + 1];
                                return nextRow.Cells[j];
                            }
                        }
                    }
                }
            }

            return "Not Found";
        }

        private Dictionary<string, string> ExtractSizes(PdfXmlMapper mapping)
        {
            Dictionary<string, string> sizeModels = new Dictionary<string, string>();

            foreach (var page in mapping.Pages)
            {
                foreach (var table in page.Tables)
                {
                    if (table.Rows.Count > 1)
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];

                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                var cell = row.Cells[j];
                                if (cell.Trim().Equals(Constants.Common.ScheduledLine))
                                {
                                    var sizeRow = table.Rows[table.Rows.Count - 1];
                                    Regex regex = new Regex("[ ]{2,}", RegexOptions.None);

                                    string sizeLine = sizeRow.Cells[3];
                                    sizeLine = sizeLine.Replace("-", "").Trim();
                                    sizeLine = regex.Replace(sizeLine, " ");
                                    var sizes = sizeLine.Split(' ');

                                    string sizeValueLine = sizeRow.Cells[4];
                                    sizeValueLine = sizeValueLine.Replace("-", "").Trim();
                                    sizeValueLine = regex.Replace(sizeValueLine, " ");
                                    var sizeValues = sizeValueLine.Split(' ');

                                    if (sizes.Length == sizeValues.Length)
                                    {
                                        for (int sizeIndex = 0; sizeIndex < sizes.Length; sizeIndex++)
                                        {
                                            sizeModels.Add(sizes[sizeIndex], sizeValues[sizeIndex]);
                                        }
                                    }
                                    else
                                    {
                                        throw new Exception("The sizes count and values count are different.");
                                    }

                                    return sizeModels;
                                }
                            }
                        }
                    }
                }
            }

            throw new Exception("Cannot find Sizes");
        }
    }
}
