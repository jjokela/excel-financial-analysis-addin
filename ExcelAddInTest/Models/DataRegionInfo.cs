using System.Collections.Generic;

namespace ExcelAddInTest.Models
{
    public class DataRegionInfo
    {
        public string Type { get; set; } = "DataRegion";
        public string Name { get; set; }
        public string Range { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public bool LikelyHasHeader { get; set; }
        public List<string> FirstRowTypes { get; set; } = new List<string>();
        public List<string> SecondRowTypes { get; set; } = new List<string>();
        public List<ColumnInfo> Columns { get; set; } = new List<ColumnInfo>();
        public List<Dictionary<string, object>> SampleData { get; set; } = new List<Dictionary<string, object>>();
    }
}
