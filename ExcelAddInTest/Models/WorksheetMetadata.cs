namespace ExcelAddInTest.Models
{
    public class WorksheetMetadata
    {
        public string WorksheetName { get; set; }
        public string UsedRange { get; set; }
        public int UsedRangeRowCount { get; set; }
        public int UsedRangeColumnCount { get; set; }
        public int UsedRangeFirstRow { get; set; }
        public int UsedRangeFirstColumn { get; set; }
    }
}
