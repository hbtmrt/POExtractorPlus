using ClosedXML.Excel;
using POExtractorPlus.ExtractionBehavior;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace POExtractorPlus.Accounts
{
    public class LaclMexico : IAccount
    {
        public Dictionary<string, ExtractedRowModel> ExtractedLines { get; set; }
        public List<string> RejectedFiles { get; set; }
        Thread exportThread;
        public string Destination { get; set; }

        public LaclMexico()
        {
            this.ExtractedLines = new Dictionary<string, ExtractedRowModel>();
            this.RejectedFiles = new List<string>();
            exportThread = new Thread(Export);
        }

        public List<string> Extract(string destination, string[] files)
        {
            this.Destination = destination;
            Dictionary<string, string> fileContents = Common.GetFileContents(files, Common.ExtractionTechnology.Itext);

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
            //ws.Cell("F1").Style.NumberFormat.Format = "dd-mmm-yy";
            ws.Column(6).Style.NumberFormat.Format = "dd-mmm-yy";
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
            string[] lines = fileXmlContent.Split('\n');

            if (lines.Length > 0)
            {
                try
                {
                    // Extract PO Number
                    model.PONumber = ExtractPONumber(lines);
                    model.SeasonCode = ExtractSeasonCode(lines);
                    model.Manufacturer = ExtractManufacturer(lines);
                    model.Material = ExtractMaterial(lines);
                    model.MaterialDescription = ExtractMaterialDescription(lines);
                    model.Plan_ExFTY = ExtractPlannedDelDate(lines);
                    model.Original_ExFTY = "";
                    model.TotalPOQty = ExtractTotalPOLineQty(lines);
                    model.POUnitPrice = ExtractPOUnitPrice(lines);
                    model.DeliveryAddress = ExtractDeliveryAddress(lines);
                    model.TransMode = ExtractTransMode(lines);
                    model.Sizes = ExtractSizes(lines);
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

        private string ExtractPONumber(string[] lines)
        {
            foreach (string line in lines)
            {
                if (line.Contains(Constants.LaclMexico.PONumber))
                {
                    var splits1 = line.Split(new string[] { Constants.LaclMexico.PONumber }, StringSplitOptions.RemoveEmptyEntries);
                    if (splits1.Count() > 0) {
                        var splits2 = splits1[0].Split(' ');

                        if (splits2.Count() > 0)
                        {
                            return splits2[1];
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractSeasonCode(string[] lines)
        {
            foreach (string line in lines)
            {
                if (line.Contains(Constants.LaclMexico.SeasonCode))
                {
                    var splits1 = line.Split(new string[] { Constants.LaclMexico.SeasonCode }, StringSplitOptions.RemoveEmptyEntries);
                    if (splits1.Count() > 0)
                    {
                        var splits2 = splits1[1].Split(' ');

                        if (splits2.Count() > 0)
                        {
                            return splits2[1];
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractManufacturer(string[] lines)
        {
            foreach (string line in lines)
            {
                if (line.Contains(Constants.LaclMexico.Manufacturer))
                {
                    var splits1 = line.Split(new string[] { Constants.LaclMexico.Manufacturer }, StringSplitOptions.RemoveEmptyEntries);
                    if (splits1.Count() > 0)
                    {
                        var splits2 = splits1[1].Split(' ');

                        if (splits2.Count() > 0)
                        {
                            return splits2[1];
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractMaterial(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.Material))
                {
                    string dataLine = lines[i + 2];
                    var splits1 = dataLine.Split(' ');
                    return splits1[0];
                }
            }

            return "";
        }

        private string ExtractMaterialDescription(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.MaterialDescription))
                {
                    string dataLine = lines[i + 2];
                    var splits1 = dataLine.Split(' ');
                    string description = string.Empty;

                    for (int j = 1; j < splits1.Length; j++)
                    {
                        if (!IsNumeric(splits1[j])) {
                            description = string.Format("{0} {1}", description, splits1[j]);
                        }
                        else
                        {
                            return description;
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractDeliveryAddress(string[] lines)
        {
            foreach (string line in lines)
            {
                if (line.Contains(Constants.LaclMexico.ShipTo))
                {
                    var splits1 = line.Split(new string[] { Constants.LaclMexico.ShipTo }, StringSplitOptions.RemoveEmptyEntries);
                    if (splits1.Count() > 0)
                    {
                        string substring = splits1[0];
                        int indexOfSellTo = substring.IndexOf(Constants.LaclMexico.SellTo);
                        string newSubstring = substring.Substring(0, indexOfSellTo);
                        return newSubstring;
                    }
                }
            }

            return "";
        }

        private string ExtractPannedExFacDate(string[] lines)
        {
            

            return "";
        }

        private string ExtractOriginalExFacDate(string[] lines)
        {
            
            return "";
        }

        private string ExtractPOUnitPrice(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.Material))
                {
                    string dataLine = lines[i + 2];
                    var splits1 = dataLine.Split(' ');

                    for (int j = 0; j < splits1.Length; j++)
                    {
                        string splitItem = splits1[j];
                        if (splitItem.Length > 8 && !IsNumeric(splitItem) && IsADate(splitItem, "dd-MMM-y"))
                        {
                            return splits1[j + 3];
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractTotalPOLineQty(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.Material))
                {
                    string dataLine = lines[i + 2];
                    var splits1 = dataLine.Split(' ');

                    for (int j = 0; j < splits1.Length; j++)
                    {
                        string splitItem = splits1[j];
                        if (splitItem.Length > 8 && !IsNumeric(splitItem) && IsADate(splitItem, "dd-MMM-y"))
                        {
                            return splits1[j-2];
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractPlannedDelDate(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.Material))
                {
                    string dataLine = lines[i + 2];
                    var splits1 = dataLine.Split(' ');

                    foreach (var splitItem in splits1)
                    {
                        if (splitItem.Length > 8 && !IsNumeric(splitItem) && IsADate(splitItem, "dd-MMM-y")) {
                            return splitItem;
                        }
                    }
                }
            }

            return "";
        }

        private bool IsNumeric(string splitItem)
        {
            bool isNumeric = false;

            if (!string.IsNullOrEmpty(splitItem))
            {
                splitItem = splitItem.Trim().Replace(",", "");

                int n;
                isNumeric = int.TryParse(splitItem, out n);
            }

            return isNumeric;
        }

        private bool IsADate(string splitItem, string dateFormat)
        {
            bool isADate = false;

            if (!string.IsNullOrEmpty(splitItem)) {
                splitItem = splitItem.Trim();

                if (string.IsNullOrEmpty(dateFormat))
                {
                    DateTime value;
                    isADate = DateTime.TryParse(splitItem, out value);
                }
                else
                {
                    DateTime value;
                    isADate = DateTime.TryParseExact(
                        splitItem, dateFormat, new CultureInfo("en-US"), DateTimeStyles.None, out value);
                }
            }

            return isADate;
        }

        private string ExtractTransMode(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.Material))
                {
                    string dataLine = lines[i + 2];
                    var splits1 = dataLine.Split(' ');

                    for (int j = 0; j < splits1.Length; j++)
                    {
                        string splitItem = splits1[j];
                        if (splitItem.Length > 8 && !IsNumeric(splitItem) && IsADate(splitItem, "dd-MMM-y"))
                        {
                            return splits1[j + 1];
                        }
                    }
                }
            }

            return "";
        }

        private Dictionary<string, string> ExtractSizes(string[] lines)
        {
            Dictionary<string, string> sizeModels = new Dictionary<string, string>();

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Contains(Constants.LaclMexico.Size) && line.Contains(Constants.LaclMexico.Qty))
                {
                    for (int j = i + 1; j < lines.Length; j++)
                    {
                        string sizeLine = lines[j];
                        var splits = sizeLine.Split(' ');
                        if (splits.Length > 4)
                        {
                            bool isASize = Common.GetPossibleSizes().Contains(splits[0].Trim());
                            bool isNumeric = IsNumeric(splits[1]);
                            if (isASize && isNumeric)
                            {
                                sizeModels.Add(splits[0], splits[1]);
                            }
                        }
                    }

                    break;
                }
            }

            return sizeModels;
        }
    }
}
