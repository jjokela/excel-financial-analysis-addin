using ExcelAddInTest.Command;
using ExcelAddInTest.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.ViewModels
{
    public class WorksheetInfoViewModel : ViewModelBase
    {
        private readonly Excel.Application _excelApp;
        private string _statusText;
        private ObservableCollection<TableInfo> _tables = new ObservableCollection<TableInfo>();
        private ObservableCollection<DataRegionInfo> _dataRegions = new ObservableCollection<DataRegionInfo>();

        public DelegateCommand CaptureImageCommand { get; }
        public DelegateCommand RefreshTablesCommand { get; }
        public DelegateCommand CopyTablesJsonCommand { get; }
        public DelegateCommand DetectDataRegionsCommand { get; }

        public WorksheetInfoViewModel(Excel.Application excelApp)
        {
            _excelApp = excelApp;
            CaptureImageCommand = new DelegateCommand(CaptureWorksheetImage);
            RefreshTablesCommand = new DelegateCommand(RefreshTables);
            CopyTablesJsonCommand = new DelegateCommand(CopyTablesJson, CanCopyTablesJson);
            DetectDataRegionsCommand = new DelegateCommand(DetectDataRegions);
        }

        public string StatusText
        {
            get => _statusText;
            set
            {
                _statusText = value;
                RaisePropertyChanged();
            }
        }

        public ObservableCollection<TableInfo> Tables
        {
            get => _tables;
            set
            {
                _tables = value;
                RaisePropertyChanged();
                CopyTablesJsonCommand.RaiseCanExecuteChanged();
            }
        }

        public ObservableCollection<DataRegionInfo> DataRegions
        {
            get => _dataRegions;
            set
            {
                _dataRegions = value;
                RaisePropertyChanged();
                CopyTablesJsonCommand.RaiseCanExecuteChanged();
            }
        }

        private bool CanCopyTablesJson(object obj) => Tables.Any() || DataRegions.Any();

        private void CopyTablesJson(object obj)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                // Get worksheet metadata
                WorksheetMetadata worksheetMetadata = null;
                if (_excelApp.ActiveSheet is Excel.Worksheet worksheet)
                {
                    var usedRange = worksheet.UsedRange;
                    worksheetMetadata = new WorksheetMetadata
                    {
                        WorksheetName = worksheet.Name,
                        UsedRange = usedRange?.Address ?? "N/A",
                        UsedRangeRowCount = usedRange?.Rows.Count ?? 0,
                        UsedRangeColumnCount = usedRange?.Columns.Count ?? 0,
                        UsedRangeFirstRow = usedRange?.Row ?? 0,
                        UsedRangeFirstColumn = usedRange?.Column ?? 0
                    };
                }

                // Combine Tables and DataRegions into a unified list
                var allData = new List<object>();

                foreach (var table in Tables)
                {
                    allData.Add(table);
                }

                foreach (var region in DataRegions)
                {
                    // Create a clean object without the debug type arrays
                    allData.Add(new
                    {
                        region.Type,
                        region.Name,
                        region.Range,
                        region.RowCount,
                        region.ColumnCount,
                        region.LikelyHasHeader,
                        region.Columns,
                        region.SampleData
                    });
                }

                // Create output with metadata wrapper
                var output = new
                {
                    Worksheet = worksheetMetadata,
                    Data = allData
                };

                var json = JsonSerializer.Serialize(output, options);
                Clipboard.SetText(json);

                var parts = new List<string>();
                if (Tables.Any()) parts.Add($"{Tables.Count} table(s)");
                if (DataRegions.Any()) parts.Add($"{DataRegions.Count} data region(s)");
                StatusText = $"JSON copied: {string.Join(" + ", parts)}.";
            }
            catch (Exception ex)
            {
                StatusText = $"Error copying JSON: {ex.Message}";
            }
        }

        private void RefreshTables(object obj)
        {
            try
            {
                Tables.Clear();

                if (!(_excelApp.ActiveSheet is Excel.Worksheet worksheet))
                {
                    StatusText = "No active worksheet found.";
                    return;
                }

                int tableCount = 0;
                int autoFilterCount = 0;

                // Get formal Tables (ListObjects)
                var listObjects = worksheet.ListObjects;
                foreach (Excel.ListObject listObject in listObjects)
                {
                    var tableInfo = new TableInfo
                    {
                        Type = "Table",
                        Name = listObject.Name,
                        Range = listObject.Range?.Address ?? "N/A",
                        RowCount = listObject.ListRows.Count,
                        ColumnCount = listObject.ListColumns.Count
                    };

                    // Get column information with data types
                    foreach (Excel.ListColumn column in listObject.ListColumns)
                    {
                        var columnInfo = new ColumnInfo
                        {
                            Name = column.Name,
                            Index = column.Index,
                            DataType = DetectColumnDataType(column.DataBodyRange)
                        };
                        tableInfo.Columns.Add(columnInfo);
                    }

                    // Extract sample data (first 5 rows)
                    ExtractSampleData(tableInfo, listObject.DataBodyRange, listObject.ListColumns);

                    Tables.Add(tableInfo);
                    tableCount++;
                }

                // Get standalone AutoFilter (not part of a Table)
                if (worksheet.AutoFilter != null && worksheet.AutoFilterMode)
                {
                    var autoFilter = worksheet.AutoFilter;
                    var filterRange = autoFilter.Range;

                    // Check if this AutoFilter is not already covered by a ListObject
                    bool isStandalone = true;
                    foreach (Excel.ListObject listObject in listObjects)
                    {
                        if (listObject.Range.Address == filterRange.Address)
                        {
                            isStandalone = false;
                            break;
                        }
                    }

                    if (isStandalone)
                    {
                        var headerRow = filterRange.Rows[1] as Excel.Range;
                        var dataRows = filterRange.Rows.Count - 1;
                        var dataRange = dataRows > 0 ? filterRange.Offset[1, 0].Resize[dataRows, filterRange.Columns.Count] : null;

                        var tableInfo = new TableInfo
                        {
                            Type = "AutoFilter",
                            Name = "AutoFilter Range",
                            Range = filterRange.Address,
                            RowCount = dataRows,
                            ColumnCount = filterRange.Columns.Count
                        };

                        // Get column names and data types from header row
                        for (int col = 1; col <= filterRange.Columns.Count; col++)
                        {
                            var headerCell = headerRow.Cells[1, col] as Excel.Range;
                            var columnDataRange = dataRows > 0 ? (filterRange.Columns[col] as Excel.Range)?.Offset[1, 0].Resize[dataRows, 1] : null;
                            var columnInfo = new ColumnInfo
                            {
                                Name = headerCell?.Value?.ToString() ?? $"Column{col}",
                                Index = col,
                                DataType = DetectColumnDataType(columnDataRange)
                            };
                            tableInfo.Columns.Add(columnInfo);
                        }

                        // Extract sample data
                        ExtractSampleDataFromRange(tableInfo, dataRange, headerRow);

                        Tables.Add(tableInfo);
                        autoFilterCount++;
                    }
                }

                if (Tables.Count == 0)
                {
                    StatusText = "No tables or filtered ranges found in the active worksheet.";
                }
                else
                {
                    var parts = new System.Collections.Generic.List<string>();
                    if (tableCount > 0) parts.Add($"{tableCount} table(s)");
                    if (autoFilterCount > 0) parts.Add($"{autoFilterCount} filtered range(s)");
                    StatusText = $"Found {string.Join(" and ", parts)}.";
                }

                CopyTablesJsonCommand.RaiseCanExecuteChanged();
            }
            catch (Exception ex)
            {
                StatusText = $"Error refreshing tables: {ex.Message}";
            }
        }

        private void DetectDataRegions(object obj)
        {
            try
            {
                DataRegions.Clear();

                if (!(_excelApp.ActiveSheet is Excel.Worksheet worksheet))
                {
                    StatusText = "No active worksheet found.";
                    return;
                }

                var usedRange = worksheet.UsedRange;
                if (usedRange == null)
                {
                    StatusText = "No used range found.";
                    return;
                }

                // Collect ranges to exclude (Excel Tables and AutoFilter ranges)
                var excludedRanges = new List<Excel.Range>();

                // Add all ListObject (Table) ranges
                foreach (Excel.ListObject listObject in worksheet.ListObjects)
                {
                    if (listObject.Range != null)
                        excludedRanges.Add(listObject.Range);
                }

                // Add AutoFilter range if present
                if (worksheet.AutoFilter != null && worksheet.AutoFilterMode)
                {
                    excludedRanges.Add(worksheet.AutoFilter.Range);
                }

                Excel.Range constantsRange = null;
                try
                {
                    // Get all cells with constants (non-formula, non-empty values)
                    constantsRange = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // No constants found
                    StatusText = "No constant data regions found.";
                    return;
                }

                if (constantsRange == null)
                {
                    StatusText = "No constant data regions found.";
                    return;
                }

                int skippedInTables = 0;
                int skippedNoHeader = 0;

                // Each Area is a contiguous block of data
                foreach (Excel.Range area in constantsRange.Areas)
                {
                    // Check if this area overlaps with any excluded range
                    bool isExcluded = false;
                    foreach (var excludedRange in excludedRanges)
                    {
                        Excel.Range intersection = _excelApp.Intersect(area, excludedRange);
                        if (intersection != null)
                        {
                            isExcluded = true;
                            skippedInTables++;
                            break;
                        }
                    }

                    if (isExcluded)
                        continue;

                    int rowCount = area.Rows.Count;
                    int colCount = area.Columns.Count;

                    var regionInfo = new DataRegionInfo
                    {
                        Name = $"Region at {area.Address}",
                        Range = area.Address,
                        RowCount = rowCount,
                        ColumnCount = colCount
                    };

                    // Read up to 6 rows in ONE COM call (1 potential header + 5 sample data rows)
                    // This gives us everything we need for analysis
                    int rowsToRead = Math.Min(rowCount, 6);
                    var sampleRange = area.Resize[rowsToRead, colCount];
                    var values = sampleRange.Value2; // Single COM call!

                    // Handle single cell vs array
                    object[,] arr = null;
                    if (values is object[,] multiCell)
                    {
                        arr = multiCell;
                    }
                    else if (rowsToRead == 1 && colCount == 1)
                    {
                        // Single cell - wrap in array for uniform handling
                        arr = new object[2, 2];
                        arr[1, 1] = values;
                    }

                    if (arr != null)
                    {
                        // Analyze first row types
                        for (int col = 1; col <= colCount; col++)
                        {
                            regionInfo.FirstRowTypes.Add(GetValueType(arr[1, col]));
                        }

                        // Analyze second row types (if exists)
                        if (rowCount >= 2)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                regionInfo.SecondRowTypes.Add(GetValueType(arr[2, col]));
                            }

                            // Determine if likely header
                            bool firstRowAllText = regionInfo.FirstRowTypes.All(t => t == "Text" || t == "Empty");
                            bool secondRowHasNonText = regionInfo.SecondRowTypes.Any(t => t != "Text" && t != "Empty");
                            regionInfo.LikelyHasHeader = firstRowAllText && secondRowHasNonText;
                        }

                        // Extract column info and sample data
                        int dataStartRow = regionInfo.LikelyHasHeader ? 2 : 1;
                        int dataRowCount = regionInfo.LikelyHasHeader ? rowCount - 1 : rowCount;
                        regionInfo.RowCount = dataRowCount; // Update to reflect actual data rows

                        // Build column info
                        for (int col = 1; col <= colCount; col++)
                        {
                            string colName;
                            string dataType;

                            if (regionInfo.LikelyHasHeader)
                            {
                                colName = arr[1, col]?.ToString() ?? $"Column{col}";
                                dataType = regionInfo.SecondRowTypes.Count >= col ? regionInfo.SecondRowTypes[col - 1] : "Unknown";
                            }
                            else
                            {
                                colName = $"Column{col}";
                                dataType = regionInfo.FirstRowTypes.Count >= col ? regionInfo.FirstRowTypes[col - 1] : "Unknown";
                            }

                            regionInfo.Columns.Add(new ColumnInfo
                            {
                                Name = colName,
                                Index = col,
                                DataType = dataType
                            });
                        }

                        // Extract sample data (up to 5 rows)
                        int sampleRows = Math.Min(5, rowsToRead - (regionInfo.LikelyHasHeader ? 1 : 0));
                        for (int row = dataStartRow; row < dataStartRow + sampleRows && row <= rowsToRead; row++)
                        {
                            var rowData = new Dictionary<string, object>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                var colName = regionInfo.Columns[col - 1].Name;
                                rowData[colName] = FormatCellValue(arr[row, col]);
                            }
                            regionInfo.SampleData.Add(rowData);
                        }
                    }

                    // Only include regions that likely have headers (indicating table structure)
                    if (!regionInfo.LikelyHasHeader)
                    {
                        skippedNoHeader++;
                        continue;
                    }

                    DataRegions.Add(regionInfo);
                }

                var skipParts = new List<string>();
                if (skippedInTables > 0) skipParts.Add($"{skippedInTables} in tables/filters");
                if (skippedNoHeader > 0) skipParts.Add($"{skippedNoHeader} without headers");
                var skippedMsg = skipParts.Any() ? $" (skipped: {string.Join(", ", skipParts)})" : "";
                StatusText = $"Found {DataRegions.Count} table-like region(s){skippedMsg}.";
                CopyTablesJsonCommand.RaiseCanExecuteChanged();
            }
            catch (Exception ex)
            {
                StatusText = $"Error detecting data regions: {ex.Message}";
            }
        }

        /// <summary>
        /// Determines the data type from a raw Value2 cell value.
        /// Pure .NET - no COM calls.
        /// </summary>
        private string GetValueType(object value)
        {
            if (value == null) return "Empty";

            // Value2 returns doubles for dates (OLE Automation date)
            // We can't distinguish date from number without NumberFormat (which requires COM)
            // So we just report the raw type here
            if (value is double) return "Number";
            if (value is bool) return "Boolean";
            if (value is string) return "Text";
            if (value is int) return "Number";

            return "Unknown";
        }

        private string DetectColumnDataType(Excel.Range dataRange)
        {
            if (dataRange == null) return "Unknown";

            try
            {
                // Sample a few cells to determine type
                int sampleSize = Math.Min(5, dataRange.Rows.Count);
                var types = new Dictionary<string, int>();

                for (int row = 1; row <= sampleSize; row++)
                {
                    var cell = dataRange.Cells[row, 1] as Excel.Range;
                    if (cell?.Value == null) continue;

                    var cellType = GetCellDataType(cell);
                    if (!types.ContainsKey(cellType))
                        types[cellType] = 0;
                    types[cellType]++;
                }

                if (types.Count == 0) return "Empty";

                // Return the most common type
                return types.OrderByDescending(x => x.Value).First().Key;
            }
            catch
            {
                return "Unknown";
            }
        }

        private string GetCellDataType(Excel.Range cell)
        {
            var value = cell.Value;

            if (value == null) return "Empty";

            // Check value type FIRST - this is the most reliable indicator
            if (value is string) return "Text";
            if (value is bool) return "Boolean";
            if (value is DateTime) return "Date";

            // For numeric values, check NumberFormat to distinguish Currency/Percentage/Date
            if (value is double || value is int || value is decimal)
            {
                var numberFormat = cell.NumberFormat?.ToString() ?? "";

                // Check for percentage
                if (numberFormat.Contains("%")) return "Percentage";

                // Check for currency symbols
                if (numberFormat.Contains("$") || numberFormat.Contains("€") || numberFormat.Contains("£") ||
                    numberFormat.Contains("¥") || numberFormat.Contains("₹"))
                    return "Currency";

                // Check for date/time patterns (but only for numeric values that could be dates)
                // Excel stores dates as numbers, so a double with date format is a date
                if (numberFormat.Contains("d") || numberFormat.Contains("m") || numberFormat.Contains("y") ||
                    numberFormat.Contains("h") || numberFormat.Contains("s"))
                    return "Date";

                return "Number";
            }

            return "Unknown";
        }

        private void ExtractSampleData(TableInfo tableInfo, Excel.Range dataRange, Excel.ListColumns columns)
        {
            if (dataRange == null) return;

            try
            {
                int sampleRows = Math.Min(5, dataRange.Rows.Count);

                for (int row = 1; row <= sampleRows; row++)
                {
                    var rowData = new Dictionary<string, object>();

                    foreach (Excel.ListColumn column in columns)
                    {
                        var cell = dataRange.Cells[row, column.Index] as Excel.Range;
                        var value = cell?.Value;
                        rowData[column.Name] = FormatCellValue(value);
                    }

                    tableInfo.SampleData.Add(rowData);
                }
            }
            catch
            {
                // Sample data extraction failed, continue without it
            }
        }

        private void ExtractSampleDataFromRange(TableInfo tableInfo, Excel.Range dataRange, Excel.Range headerRow)
        {
            if (dataRange == null) return;

            try
            {
                int sampleRows = Math.Min(5, dataRange.Rows.Count);
                int columnCount = dataRange.Columns.Count;

                for (int row = 1; row <= sampleRows; row++)
                {
                    var rowData = new Dictionary<string, object>();

                    for (int col = 1; col <= columnCount; col++)
                    {
                        var headerCell = headerRow.Cells[1, col] as Excel.Range;
                        var columnName = headerCell?.Value?.ToString() ?? $"Column{col}";
                        var cell = dataRange.Cells[row, col] as Excel.Range;
                        var value = cell?.Value;
                        rowData[columnName] = FormatCellValue(value);
                    }

                    tableInfo.SampleData.Add(rowData);
                }
            }
            catch
            {
                // Sample data extraction failed, continue without it
            }
        }

        private object FormatCellValue(object value)
        {
            if (value == null) return null;
            if (value is DateTime dt) return dt.ToString("yyyy-MM-dd");
            if (value is double d) return Math.Round(d, 4);
            return value;
        }

        private string GetTableStyleName(Excel.ListObject listObject)
        {
            try
            {
                var style = listObject.TableStyle;
                if (style == null)
                    return "None";

                // TableStyle can be a TableStyle object or a string
                if (style is Excel.TableStyle tableStyle)
                    return tableStyle.Name;

                return style.ToString();
            }
            catch
            {
                return "Unknown";
            }
        }

        private void CaptureWorksheetImage(object obj)
        {
            try
            {
                if (!(_excelApp.ActiveSheet is Excel.Worksheet worksheet))
                {
                    StatusText = "No active worksheet found.";
                    return;
                }

                var usedRange = worksheet.UsedRange;
                if (usedRange == null || usedRange.Cells.Count == 0)
                {
                    StatusText = "Worksheet has no data to capture.";
                    return;
                }

                // Copy range as picture to clipboard
                usedRange.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);

                // Get image from clipboard
                var image = Clipboard.GetImage();
                if (image == null)
                {
                    StatusText = "Failed to capture image from clipboard.";
                    return;
                }

                // Determine save folder
                var folder = Properties.Settings.Default.ImageSaveFolder;
                if (string.IsNullOrEmpty(folder))
                {
                    folder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                }

                // Ensure folder exists
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                // Generate filename with timestamp
                var filename = $"worksheet_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                var fullPath = Path.Combine(folder, filename);

                // Save image
                image.Save(fullPath, ImageFormat.Png);
                image.Dispose();

                StatusText = $"Image saved to: {fullPath}";
            }
            catch (Exception ex)
            {
                StatusText = $"Error capturing image: {ex.Message}";
            }
        }
    }
}
