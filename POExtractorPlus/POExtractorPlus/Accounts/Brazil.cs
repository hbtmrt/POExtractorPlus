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
    public class Brazil : IAccount
    {
        public Dictionary<string, ExtractedRowModel> ExtractedLines { get; set; }
        public List<string> RejectedFiles { get; set; }
        Thread exportThread;
        public string Destination { get; set; }

        public Brazil()
        {
            this.ExtractedLines = new Dictionary<string, ExtractedRowModel>();
            this.RejectedFiles = new List<string>();
            exportThread = new Thread(Export);
        }

        public List<string> Extract(string destination, string[] files)
        {
            this.Destination = destination;
            Dictionary<string, string> fileContents = Common.GetFileContents(files, Common.ExtractionTechnology.Sautin);

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
                    model.Original_ExFTY = "";
                    model.TotalPOQty = ExtractTotalPOQty(mapping);
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

            return "";
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

            return "";
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

            return "";
        }

        // Extract Product
        private string ExtractMaterial(PdfXmlMapper mapping)
        {
            if (mapping.Pages.Count > 0 &&
                mapping.Pages[0].Tables.Count > 1 &&
                mapping.Pages[0].Tables[0].Rows.Count > 0 &&
                mapping.Pages[0].Tables[0].Rows[0].Cells.Count > 0)
            {
                return mapping.Pages[0].Tables[0].Rows[0].Cells[0];
            }

            return "";
        }

        private string ExtractMaterialDescription(PdfXmlMapper mapping)
        {
            if (mapping.Pages.Count > 0 &&
                mapping.Pages[0].Tables.Count > 1 &&
                mapping.Pages[0].Tables[0].Rows.Count > 0 &&
                mapping.Pages[0].Tables[0].Rows[0].Cells.Count > 1)
            {
                return mapping.Pages[0].Tables[0].Rows[0].Cells[2];
            }

            return "";
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
                            if (cell.Contains(Constants.Common.DeliveryAddress))
                            {
                                return cell.Replace(Constants.Common.DeliveryAddress, "").Trim();
                            }
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractPOUnitPrice(PdfXmlMapper mapping)
        {
            if (mapping.Pages.Count > 0 &&
                mapping.Pages[0].Tables.Count > 1 &&
                mapping.Pages[0].Tables[0].Rows.Count > 0 &&
                mapping.Pages[0].Tables[0].Rows[0].Cells.Count > 19)
            {
                return mapping.Pages[0].Tables[0].Rows[0].Cells[20];
            }

            return "";
        }

        private string ExtractTotalPOQty(PdfXmlMapper mapping)
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
                            if (cell.Equals(Constants.Brazil.TotalPOQty))
                            {
                                string poQty = row.Cells[i + 1];
                                return poQty;
                            }
                        }
                    }
                }
            }

            return "";
        }

        private string ExtractPlannedDelDate(PdfXmlMapper mapping)
        {
            if (mapping.Pages.Count > 0 &&
                mapping.Pages[0].Tables.Count > 1 &&
                mapping.Pages[0].Tables[0].Rows.Count > 0 &&
                mapping.Pages[0].Tables[0].Rows[0].Cells.Count > 13)
            {
                return mapping.Pages[0].Tables[0].Rows[0].Cells[14];
            }

            return "";
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
                            if (cell.Contains(Constants.Brazil.TransMode))
                            {
                                return cell.Replace(Constants.Brazil.TransMode, "").Trim();
                            }
                        }
                    }
                }
            }

            return "";
        }

        private Dictionary<string, string> ExtractSizes(PdfXmlMapper mapping)
        {
            Dictionary<string, string> sizeModels = new Dictionary<string, string>();

            //the first line is not a column, it is a line, so we have to extract it in ugly way. :D
            if (mapping.Pages.Count > 0 &&
                mapping.Pages[0].Tables.Count > 1 &&
                mapping.Pages[0].Tables[1].Rows.Count > 0 &&
                mapping.Pages[0].Tables[1].Rows[mapping.Pages[0].Tables[1].Rows.Count - 1].Cells.Count > 0)
            {
                string firstSizeRow = mapping.Pages[0].Tables[1].Rows[mapping.Pages[0].Tables[1].Rows.Count - 1].Cells[0];
                string[] splits = firstSizeRow.Split(' ');
                string sizeType = splits[splits.Length - 7];
                string sizeValue = splits[splits.Length - 3];
                sizeModels.Add(sizeType, sizeValue);
            }

            //extract other tables.
            if (mapping.Pages.Count > 0 &&
                mapping.Pages[0].Tables.Count > 1 &&
                mapping.Pages[0].Tables[0].Rows.Count > 0)
            {
                foreach (var row in mapping.Pages[0].Tables[0].Rows)
                {
                    if (row.Cells.Count > 20)
                    {
                        string sizeType = row.Cells[10];
                        string sizeValue = row.Cells[18];
                        sizeModels.Add(sizeType, sizeValue);
                    }
                }
            }

            return sizeModels;
        }
    }
}
