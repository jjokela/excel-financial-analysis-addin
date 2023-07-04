using ExcelAddInTest.ViewModels;
using Moq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInTest.Test;

public class FinancialStatementAnalysisViewModelTests
{
    private Mock<Excel.Application> _mockExcelApp = null!;
    private Mock<Excel.Range> _mockRange = null!;

    [SetUp]
    public void Setup()
    {
        _mockExcelApp = new Mock<Excel.Application>();
        _mockRange = new Mock<Excel.Range>();
    }

    [Test]
    public void GetText_ReturnsExpectedResult()
    {
        var viewModel = new FinancialStatementAnalysisViewModel(_mockExcelApp.Object);
        _mockRange.Setup(m => m.Rows.Count).Returns(2);
        _mockRange.Setup(m => m.Columns.Count).Returns(2);
        _mockExcelApp.Setup(m => m.Selection).Returns(_mockRange.Object);

        var excelValues = new[,] { { "Test1", "Test2" }, { "Test3", "Test4" } };

        for (var row = 0; row < 2; row++)
        {
            for (var col = 0; col < 2; col++)
            {
                var mockCell = new Mock<IExcelCellValue>();
                mockCell.Setup(x => x.Value).Returns(excelValues[row, col]);
                _mockRange.Setup(m => m.Cells[row + 1, col + 1]).Returns(mockCell.Object);
            }
        }

        viewModel.GetText(null);

        var expected = "Test1;Test2\r\nTest3;Test4\r\n";
        Assert.That(viewModel.AnalysisText, Is.EqualTo(expected));
    }
}

public interface IExcelCellValue
{
    string Value { get; set; }
}