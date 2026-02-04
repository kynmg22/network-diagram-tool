using System.Collections.Generic;

namespace NetworkDiagramApp
{
    public class NetworkNode
    {
        public string ID { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string IP { get; set; } = string.Empty;
        public int? VLAN { get; set; }
        public string Note { get; set; } = string.Empty;
        public List<string> Parents { get; set; } = new();
        public int ExcelRow { get; set; }

        public bool IsONU()
        {
            return ID.Trim().ToUpper() == "ONU" || Name.Trim().ToUpper() == "ONU";
        }
    }
}
