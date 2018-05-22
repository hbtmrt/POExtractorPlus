using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace POExtractorPlus
{
    [XmlRoot("document")]
    public class PdfXmlMapper
    {
        [XmlElement("page")]
        public List<PageNode> Pages { get; set; }
    }

    public class PageNode
    {
        [XmlElement("table")]
        public List<TableNode> Tables { get; set; }
    }

    public class TableNode
    {
        [XmlElement("row")]
        public List<RowNode> Rows { get; set; }
    }

    public class RowNode
    {
        [XmlElement("cell")]
        public List<string> Cells { get; set; }
    }

    public class CellNode
    {
    }
}
