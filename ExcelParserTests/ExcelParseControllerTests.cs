using ExcelParser;
using ExcelParser.Dtos;
using ExcelParser.Models;
using ExcelParser.Services.Interfaces;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NSubstitute;

namespace ExcelParserTests;

[TestFixture]
public class ExcelParseControllerTests
{
    private ExcelParseController _controller;
    private IExcelService _mockExcelService;
    private IMappingService _mockMappingService;

    [SetUp]
    public void Setup()
    {
        _mockExcelService = Substitute.For<IExcelService>();
        _mockMappingService = Substitute.For<IMappingService>();
        _controller = new ExcelParseController(_mockExcelService, _mockMappingService);
    }

    [TearDown]
    public void TearDown()
    {
        _controller.Dispose();
    }

    [TestFixture]
    public class ParseExcelTests : ExcelParseControllerTests
    {
        [Test]
        public async Task ParseExcel_WithNullFile_ReturnsBadRequest()
        {
            // Act
            var result = await _controller.ParseExcel(null);

            // Assert
            var badRequestResult = result as BadRequestObjectResult;
            Assert.That(badRequestResult, Is.Not.Null);
            Assert.That(badRequestResult.Value, Is.EqualTo("File is required."));
        }

        [Test]
        public async Task ParseExcel_WithEmptyFile_ReturnsBadRequest()
        {
            // Arrange
            var mockFile = CreateMockFile("test.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 0);

            // Act
            var result = await _controller.ParseExcel(mockFile);

            // Assert
            var badRequestResult = result as BadRequestObjectResult;
            Assert.That(badRequestResult, Is.Not.Null);
            Assert.That(badRequestResult.Value, Is.EqualTo("File is required."));
        }

        [Test]
        public async Task ParseExcel_WithInvalidContentType_ReturnsBadRequest()
        {
            // Arrange
            var mockFile = CreateMockFile("test.txt", "text/plain", 100);

            // Act
            var result = await _controller.ParseExcel(mockFile);

            // Assert
            var badRequestResult = result as BadRequestObjectResult;
            Assert.That(badRequestResult, Is.Not.Null);
            Assert.That(badRequestResult.Value, Is.EqualTo("Invalid content type. Please upload a valid Excel file."));
        }

        [Test]
        public async Task ParseExcel_WithValidXlsxFile_ReturnsOkResult()
        {
            // Arrange
            var mockFile = CreateMockFile("test.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 100);
            var mockDynamicData = CreateMockDynamicData();
            
            _mockExcelService.ParseAsync(Arg.Any<Stream>()).Returns(mockDynamicData);

            // Act
            var result = await _controller.ParseExcel(mockFile);

            // Assert
            var okResult = result as OkObjectResult;
            Assert.That(okResult, Is.Not.Null);
            
            var response = okResult.Value;
            Assert.That(response, Is.Not.Null);
            
            // Verify service was called
            await _mockExcelService.Received(1).ParseAsync(Arg.Any<Stream>());
        }

        [Test]
        public async Task ParseExcel_WithValidXlsFile_ReturnsOkResult()
        {
            // Arrange
            var mockFile = CreateMockFile("test.xls", "application/vnd.ms-excel", 100);
            var mockDynamicData = CreateMockDynamicData();
            
            _mockExcelService.ParseAsync(Arg.Any<Stream>()).Returns(mockDynamicData);

            // Act
            var result = await _controller.ParseExcel(mockFile);

            // Assert
            var okResult = result as OkObjectResult;
            Assert.That(okResult, Is.Not.Null);
            
            // Verify service was called
            await _mockExcelService.Received(1).ParseAsync(Arg.Any<Stream>());
        }
    }

    [TestFixture]
    public class ParseAsPersonsTests : ExcelParseControllerTests
    {
        [Test]
        public async Task ParseAsPersons_WithNullFile_ReturnsBadRequest()
        {
            // Act
            var result = await _controller.ParseAsPersons(null);

            // Assert
            var badRequestResult = result as BadRequestObjectResult;
            Assert.That(badRequestResult, Is.Not.Null);
            Assert.That(badRequestResult.Value, Is.EqualTo("File is required."));
        }

        [Test]
        public async Task ParseAsPersons_WithEmptyFile_ReturnsBadRequest()
        {
            // Arrange
            var mockFile = CreateMockFile("test.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 0);

            // Act
            var result = await _controller.ParseAsPersons(mockFile);

            // Assert
            var badRequestResult = result as BadRequestObjectResult;
            Assert.That(badRequestResult, Is.Not.Null);
            Assert.That(badRequestResult.Value, Is.EqualTo("File is required."));
        }

        [Test]
        public async Task ParseAsPersons_WithInvalidContentType_ReturnsBadRequest()
        {
            // Arrange
            var mockFile = CreateMockFile("test.txt", "text/plain", 100);

            // Act
            var result = await _controller.ParseAsPersons(mockFile);

            // Assert
            var badRequestResult = result as BadRequestObjectResult;
            Assert.That(badRequestResult, Is.Not.Null);
            Assert.That(badRequestResult.Value, Is.EqualTo("Invalid content type. Please upload a valid Excel file."));
        }

        [Test]
        public async Task ParseAsPersons_WithValidFile_ReturnsOkResult()
        {
            // Arrange
            var mockFile = CreateMockFile("test.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 100);
            var mockDynamicData = CreateMockDynamicData();
            var mockPersonDtos = CreateMockPersonDtos();
            
            _mockExcelService.ParseAsync(Arg.Any<Stream>()).Returns(mockDynamicData);
            
            _mockMappingService.MapToDto<PersonDto>(mockDynamicData).Returns(mockPersonDtos);

            // Act
            var result = await _controller.ParseAsPersons(mockFile);

            // Assert
            var okResult = result as OkObjectResult;
            Assert.That(okResult, Is.Not.Null);
            
            // Verify both services were called
            await _mockExcelService.Received(1).ParseAsync(Arg.Any<Stream>());
            _mockMappingService.Received(1).MapToDto<PersonDto>(mockDynamicData);
        }

        [Test]
        public async Task ParseAsPersons_WithValidFile_ReturnsCorrectResponseStructure()
        {
            // Arrange
            var mockFile = CreateMockFile("test.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 100);
            var mockDynamicData = CreateMockDynamicData();
            var mockPersonDtos = CreateMockPersonDtos();
            
            _mockExcelService.ParseAsync(Arg.Any<Stream>()).Returns(mockDynamicData);
            
            _mockMappingService.MapToDto<PersonDto>(mockDynamicData).Returns(mockPersonDtos);

            // Act
            var result = await _controller.ParseAsPersons(mockFile);

            // Assert
            var okResult = result as OkObjectResult;
            Assert.That(okResult, Is.Not.Null);
            
            var response = okResult.Value;
            Assert.That(response, Is.Not.Null);
            
            // Check response structure using reflection
            var responseType = response.GetType();
            var totalRowsProperty = responseType.GetProperty("TotalRows");
            var filteredFromProperty = responseType.GetProperty("FilteredFrom");
            var dataProperty = responseType.GetProperty("Data");
            Assert.Multiple(() =>
            {
                Assert.That(totalRowsProperty, Is.Not.Null);
                Assert.That(filteredFromProperty, Is.Not.Null);
                Assert.That(dataProperty, Is.Not.Null);
            });

            Assert.Multiple(() =>
            {
                Assert.That(totalRowsProperty.GetValue(response), Is.EqualTo(1));
                Assert.That(filteredFromProperty.GetValue(response), Is.EqualTo(2));
            });
        }
    }

    private static IFormFile CreateMockFile(string fileName, string contentType, long length)
    {
        var mockFile = Substitute.For<IFormFile>();
        mockFile.FileName.Returns(fileName);
        mockFile.ContentType.Returns(contentType);
        mockFile.Length.Returns(length);
        
        mockFile.CopyToAsync(Arg.Any<Stream>(), Arg.Any<CancellationToken>()).Returns(Task.CompletedTask);
        
        return mockFile;
    }

    private static List<DynamicExcelData> CreateMockDynamicData()
    {
        var data1 = new DynamicExcelData();
        data1.SetValue("Name", "John Doe");
        data1.SetValue("Email", "john@example.com");
        data1.SetValue("Age", 30);

        var data2 = new DynamicExcelData();
        data2.SetValue("Name", "Jane Smith");
        data2.SetValue("Email", "jane@example.com");
        data2.SetValue("Age", 25);

        return [data1, data2];
    }

    private static List<PersonDto> CreateMockPersonDtos()
    {
        return
        [
            new PersonDto
            {
                Name = "John Doe",
                Email = "john@example.com",
                Age = 30
            }
        ];
    }
} 