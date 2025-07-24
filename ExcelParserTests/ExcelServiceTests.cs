using System.Text;
using ExcelParser.Services;
using Microsoft.Extensions.Logging;
using NSubstitute;
using ClosedXML.Excel;

namespace ExcelParserTests;

[TestFixture]
public class ExcelServiceTests
{
    private ExcelService _excelService;
    private ILogger<ExcelService> _mockLogger;

    [SetUp]
    public void Setup()
    {
        _mockLogger = Substitute.For<ILogger<ExcelService>>();
        _excelService = new ExcelService(_mockLogger);
    }

    [Test]
    public async Task ParseAsync_WithValidExcelData_ReturnsCorrectData()
    {
        // Arrange
        var excelBytes = CreateSimpleExcelBytes();
        using var stream = new MemoryStream(excelBytes);

        // Act
        var result = await _excelService.ParseAsync(stream);

        // Assert
        Assert.That(result, Is.Not.Null);
        Assert.That(result, Has.Count.EqualTo(2));
        Assert.Multiple(() =>
        {

            // First row
            Assert.That(result[0].Data["Name"], Is.EqualTo("John Doe"));
            Assert.That(result[0].Data["Age"], Is.EqualTo(30.0)); // Excel returns numbers as double
            Assert.That(result[0].Data["Email"], Is.EqualTo("john@example.com"));

            // Second row
            Assert.That(result[1].Data["Name"], Is.EqualTo("Jane Smith"));
            Assert.That(result[1].Data["Age"], Is.EqualTo(25.0)); // Excel returns numbers as double
            Assert.That(result[1].Data["Email"], Is.EqualTo("jane@example.com"));
        });
    }

    [Test]
    public async Task ParseAsync_WithEmptyRows_SkipsEmptyRows()
    {
        // Arrange
        var excelBytes = CreateExcelWithEmptyRows();
        using var stream = new MemoryStream(excelBytes);

        // Act
        var result = await _excelService.ParseAsync(stream);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1)); // Only one non-empty row
        Assert.That(result[0].Data["Name"], Is.EqualTo("John Doe"));
    }

    [Test]
    public async Task ParseAsync_WithMixedDataTypes_PreservesDataTypes()
    {
        // Arrange
        var excelBytes = CreateMixedDataTypeExcel();
        using var stream = new MemoryStream(excelBytes);

        // Act
        var result = await _excelService.ParseAsync(stream);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        Assert.Multiple(() =>
        {
            Assert.That(result[0].Data["Name"], Is.TypeOf<string>());
            Assert.That(result[0].Data["Age"], Is.TypeOf<double>());
            Assert.That(result[0].Data["IsActive"], Is.TypeOf<bool>());
            Assert.That(result[0].Data["StartDate"], Is.TypeOf<DateTime>());
        });
    }

    [Test]
    public async Task ParseAsync_WithNullValues_HandlesNullsCorrectly()
    {
        // Arrange
        var excelBytes = CreateExcelWithNulls();
        using var stream = new MemoryStream(excelBytes);

        // Act
        var result = await _excelService.ParseAsync(stream);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        Assert.Multiple(() =>
        {
            Assert.That(result[0].Data.ContainsKey("Name"), Is.True);
            Assert.That(result[0].Data.ContainsKey("Email"), Is.False); // Null/empty values are not included
        });
    }

    [Test]
    public async Task ParseAsync_WithEmptyHeaders_SkipsEmptyHeaderColumns()
    {
        // Arrange
        var excelBytes = CreateExcelWithEmptyHeaders();
        using var stream = new MemoryStream(excelBytes);

        // Act
        var result = await _excelService.ParseAsync(stream);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        Assert.Multiple(() =>
        {
            Assert.That(result[0].Data.ContainsKey("Name"), Is.True);
            Assert.That(result[0].Data, Has.Count.EqualTo(1)); // Only one valid header
        });
    }

    [Test]
    public async Task ParseAsync_WithNonZeroStreamPosition_ResetsStreamPosition()
    {
        // Arrange
        var excelBytes = CreateSimpleExcelBytes();
        using var stream = new MemoryStream(excelBytes);
        stream.Position = 50; // Move stream position

        // Act
        var result = await _excelService.ParseAsync(stream);

        // Assert
        Assert.That(result, Is.Not.Null);
        Assert.That(result, Is.Not.Empty);
    }

    [Test]
    public async Task ParseAsync_LogsInformation_WhenParsingCompletes()
    {
        // Arrange
        var excelBytes = CreateSimpleExcelBytes();
        using var stream = new MemoryStream(excelBytes);

        // Act
        await _excelService.ParseAsync(stream);

        // Assert
        _mockLogger.Received(1).Log(
            LogLevel.Information,
            Arg.Any<EventId>(),
            Arg.Is<object>(v => v.ToString()!.Contains("Starting dynamic Excel parsing")),
            Arg.Any<Exception>(),
            Arg.Any<Func<object, Exception?, string>>());
    }

    private static byte[] CreateSimpleExcelBytes()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        
        // Headers
        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Age";
        worksheet.Cell(1, 3).Value = "Email";
        
        // Data
        worksheet.Cell(2, 1).Value = "John Doe";
        worksheet.Cell(2, 2).Value = 30;
        worksheet.Cell(2, 3).Value = "john@example.com";
        
        worksheet.Cell(3, 1).Value = "Jane Smith";
        worksheet.Cell(3, 2).Value = 25;
        worksheet.Cell(3, 3).Value = "jane@example.com";
        
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static byte[] CreateExcelWithEmptyRows()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        
        // Headers
        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Age";
        worksheet.Cell(1, 3).Value = "Email";
        
        // Data
        worksheet.Cell(2, 1).Value = "John Doe";
        worksheet.Cell(2, 2).Value = 30;
        worksheet.Cell(2, 3).Value = "john@example.com";
        
        // Empty row (row 3 is left empty)
        // Row 4 is also left empty
        
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static byte[] CreateMixedDataTypeExcel()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        
        // Headers
        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Age";
        worksheet.Cell(1, 3).Value = "IsActive";
        worksheet.Cell(1, 4).Value = "StartDate";
        
        // Data with mixed types
        worksheet.Cell(2, 1).Value = "John Doe";
        worksheet.Cell(2, 2).Value = 30;
        worksheet.Cell(2, 3).Value = true;
        worksheet.Cell(2, 4).Value = new DateTime(2023, 1, 15);
        
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static byte[] CreateExcelWithNulls()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        
        // Headers
        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Email";
        
        // Data with null/empty values
        worksheet.Cell(2, 1).Value = "John Doe";
        // Cell 2,2 (Email) is left empty
        
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static byte[] CreateExcelWithEmptyHeaders()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");
        
        // Headers - some empty
        worksheet.Cell(1, 1).Value = "Name";
        // Cell 1,2 is left empty (empty header)
        // Cell 1,3 is left empty (empty header)
        
        // Data
        worksheet.Cell(2, 1).Value = "John Doe";
        // Cells 2,2 and 2,3 correspond to empty headers
        
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }
} 