using ExcelParser.Attributes;
using ExcelParser.Models;
using ExcelParser.Services;
using Microsoft.Extensions.Logging;
using NSubstitute;

namespace ExcelParserTests;

[TestFixture]
public class AttributeMappingServiceTests
{
    [SetUp]
    public void Setup()
    {
        _mockLogger = Substitute.For<ILogger<AttributeMappingService>>();
        _mappingService = new AttributeMappingService(_mockLogger);
    }

    private AttributeMappingService _mappingService;
    private ILogger<AttributeMappingService> _mockLogger;

    [Test]
    public void MapToDto_WithValidData_MapsCorrectly()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" },
                { "Email", "john@example.com" },
                { "Age", 30 }
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Is.Not.Null);
        Assert.That(result, Has.Count.EqualTo(1));
        using (Assert.EnterMultipleScope())
        {
            Assert.That(result[0].Name, Is.EqualTo("John Doe"));
            Assert.That(result[0].Email, Is.EqualTo("john@example.com"));
            Assert.That(result[0].Age, Is.EqualTo(30));
        }
    }

    [Test]
    public void MapToDto_WithCaseInsensitiveColumnNames_MapsCorrectly()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "name", "John Doe" }, // lowercase
                { "EMAIL", "john@example.com" }, // uppercase
                { "Age", 30 }
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        using (Assert.EnterMultipleScope())
        {
            Assert.That(result[0].Name, Is.EqualTo("John Doe"));
            Assert.That(result[0].Email, Is.EqualTo("john@example.com"));
            Assert.That(result[0].Age, Is.EqualTo(30));
        }
    }

    [Test]
    public void MapToDto_WithFilterNullOrEmpty_FiltersEmptyValues()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" },
                { "Email", "" }, // Empty string should be filtered
                { "Age", 30 }
            }),
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "Jane Smith" },
                { "Email", "jane@example.com" },
                { "Age", 25 }
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1)); // Only Jane Smith should be included
        Assert.That(result[0].Name, Is.EqualTo("Jane Smith"));
    }

    [Test]
    public void MapToDto_WithRequiredField_SkipsRecordsWithMissingRequiredData()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Email", "john@example.com" }, // Missing required Name
                { "Age", 30 }
            }),
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "Jane Smith" },
                { "Email", "jane@example.com" },
                { "Age", 25 }
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1)); // Only Jane Smith should be included
        Assert.That(result[0].Name, Is.EqualTo("Jane Smith"));
    }

    [Test]
    public void MapToDto_WithTypeConversion_ConvertsTypesCorrectly()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" },
                { "Email", "john@example.com" },
                { "Age", "30" } // String that should convert to int
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        Assert.That(result[0].Age, Is.EqualTo(30));
        Assert.That(result[0].Age, Is.TypeOf<int>());
    }

    [Test]
    public void MapToDto_WithInvalidTypeConversion_SkipsProperty()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" },
                { "Email", "john@example.com" },
                { "Age", "invalid_number" } // Cannot convert to int
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        using (Assert.EnterMultipleScope())
        {
            Assert.That(result[0].Name, Is.EqualTo("John Doe"));
            Assert.That(result[0].Age, Is.EqualTo(0)); // Default value for int
        }
    }

    [Test]
    public void MapToDto_WithNoAttributedProperties_ReturnsEmptyList()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" },
                { "Email", "john@example.com" }
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestDtoWithoutAttributes>(dynamicData);

        // Assert
        Assert.That(result, Is.Not.Null);
        Assert.That(result, Is.Empty);
    }

    [Test]
    public void MapToDto_WithMissingColumns_SkipsUnmappedProperties()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" }
                // Missing Email and Age columns
            })
        };

        // Act
        var result = _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        Assert.That(result, Has.Count.EqualTo(1));
        using (Assert.EnterMultipleScope())
        {
            Assert.That(result[0].Name, Is.EqualTo("John Doe"));
            Assert.That(result[0].Email, Is.EqualTo(string.Empty)); // Default value
            Assert.That(result[0].Age, Is.EqualTo(0)); // Default value
        }
    }

    [Test]
    public void MapToDto_LogsCorrectInformation()
    {
        // Arrange
        var dynamicData = new List<DynamicExcelData>
        {
            CreateDynamicData(new Dictionary<string, object>
            {
                { "Name", "John Doe" },
                { "Email", "john@example.com" },
                { "Age", 30 }
            })
        };

        // Act
        _mappingService.MapToDto<TestPersonDto>(dynamicData);

        // Assert
        _mockLogger.Received().Log(
            LogLevel.Information,
            Arg.Any<EventId>(),
            Arg.Is<object>(v => v.ToString()!.Contains("Mapping") && v.ToString()!.Contains("TestPersonDto")),
            Arg.Any<Exception?>(),
            Arg.Any<Func<object, Exception?, string>>());
    }

    private static DynamicExcelData CreateDynamicData(Dictionary<string, object> data)
    {
        var dynamicData = new DynamicExcelData();
        foreach (var kvp in data) dynamicData.SetValue(kvp.Key, kvp.Value);
        return dynamicData;
    }

    // Test DTOs
    private class TestPersonDto
    {
        [ExcelColumn("Name", IsRequired = true, FilterNullOrEmpty = true)]
        public string Name { get; set; } = string.Empty;

        [ExcelColumn("Email", FilterNullOrEmpty = true)]
        public string Email { get; set; } = string.Empty;

        [ExcelColumn("Age", FilterNullOrEmpty = true)]
        public int Age { get; set; }
    }

    private class TestDtoWithoutAttributes
    {
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
    }
}