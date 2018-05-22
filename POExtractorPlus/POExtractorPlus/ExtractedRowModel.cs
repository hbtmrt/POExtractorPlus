using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POExtractorPlus
{
    public class ExtractedRowModel
    {
        public string AccountType { get; set; }
        public string PONumber { get; set; }
        public string Manufacturer { get; set; }
        public string SeasonCode { get; set; }
        public string Material { get; set; }
        public string MaterialDescription { get; set; }
        public string Plan_ExFTY { get; set; }
        public string Original_ExFTY { get; set; }
        public string TotalPOQty { get; set; }
        public string POUnitPrice { get; set; }
        public string DeliveryAddress { get; set; }
        public string TransMode { get; set; }
        public string SchedQty { get; set; }
        public Dictionary<string, string> Sizes { get; set; }
        public bool HasErrors { get; set; }
        public string Error { get; set; }
    }
}
