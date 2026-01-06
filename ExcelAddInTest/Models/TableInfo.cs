using System.Collections.Generic;

namespace ExcelAddInTest.Models
{
    public class TableInfo
    {
        public string Type { get; set; } = "Table";
        public string Name { get; set; }
        public string Range { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public List<ColumnInfo> Columns { get; set; } = new List<ColumnInfo>();
        public List<Dictionary<string, object>> SampleData { get; set; } = new List<Dictionary<string, object>>();
    }
}
