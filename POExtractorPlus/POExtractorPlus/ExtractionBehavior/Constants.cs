using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POExtractorPlus.ExtractionBehavior
{
    public class Constants
    {
        public static class Common
        {
            public const string PONumber = "PO NUMBER";
            public const string SeasonCode = "SEASON CODE";
            public const string Manufacturer = "MANUFACTURER";
            public const string Material = "Material";
            public const string MaterialDescription = "Material Description";
            public const string CompanyCode = "COMPANY CODE";
            public const string PannedExFacDate = "Planned EX-Fac Date";
            public const string OriginalExFacDate = "Original Ex- facDate.";
            public const string DeliveryAddress = "Delivery Address";
            public const string POUnitPrice = "PO Unit Price";
            public const string TotalPOLineQty = "Total PO Line Qty";
            public const string PlannedDelDate = "Planned Del. Date";
            public const string TransMode = "Trans. Mode";
            public const string ScheduledLine = "Schedule Line";
            public const string Product = "Product";
            public const string ProductDescription = "Product Description";
        }

        public static class Brazil
        {
            public const string TotalPOQty = "TOTAL PO QUANTITY";
            public const string TotalPOLineQty = "Total PO Qty";
            public const string TransMode = "TRANS MODE";
        }

        public static class LaclMexico
        {
            public const string PONumber = "PO:";
            public const string SeasonCode = "SEASON:";
            public const string Manufacturer = "Plant Code:";
            public const string Material = "PRODUCT CODE";
            public const string ShipTo = "SHIP TO:";
            public const string SellTo = "SELL TO:";
            public const string Size = "SIZE";
            public const string Qty = "QTY";
        }
    }
}
