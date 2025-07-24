# ExcelParser

A flexible and powerful ASP.NET Core 9.0 Web API for parsing Excel files with both dynamic and strongly-typed mapping
capabilities.

## 🚀 Features

- **Dynamic Excel Parsing**: Parse any Excel file structure using column headers as property names
- **Attribute-Based Mapping**: Map Excel data to strongly-typed DTOs with custom configuration
- **Flexible Data Filtering**: Filter null/empty values and enforce required fields
- **Type Safety**: Automatic type conversion with error handling
- **Good Architecture**: Service-oriented design with dependency injection
- **Comprehensive Logging**: Detailed logging for debugging and monitoring
- **Swagger Documentation**: Built-in API documentation and testing interface

## 🏗️ Architecture & Design Choices

### Core Components

1. **Dynamic Data Structure**: Uses `DynamicExcelData` model with a flexible `Dictionary<string, object>` to handle any
   Excel structure
2. **Attribute-Based Mapping**: Custom `ExcelColumnAttribute` for configuring column mapping, validation, and filtering
3. **Service Layer**: Clean separation of concerns with dedicated services for parsing and mapping
4. **Interface-Driven Design**: All services implement interfaces for better testability and maintainability

### Key Design Decisions

- **ExcelDataReader Library**: Chosen for robust support of both .xls and .xlsx formats
- **Stream-Based Processing**: Memory-efficient file handling using streams
- **Case-Insensitive Matching**: Column name matching ignores case for better user experience
- **Fail-Safe Parsing**: Continues processing even if individual rows fail, with comprehensive logging
- **Type Conversion**: Automatic type conversion with graceful error handling
- **Comprehensive Testing**: Uses real Excel files (via ClosedXML) for authentic testing scenarios

## 📋 Prerequisites

- [.NET 9.0 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
- A compatible IDE (Visual Studio, VS Code, Rider, etc.)

## 🛠️ Setup Instructions

### 1. Clone and Navigate

```bash
git clone <repository-url>
cd ExcelParser
```

### 2. Restore Dependencies

```bash
dotnet restore
```

### 3. Build the Project

```bash
dotnet build
```

### 4. Run the Application

```bash
dotnet run --project ExcelParser
```

The API will be available at:

- **HTTP**: `http://localhost:5001`
- **HTTPS**: `https://localhost:5000`
- **Swagger UI**: `https://localhost:5000/swagger` (Development environment only)

## 📚 API Endpoints

### 1. Dynamic Excel Parsing

**POST** `/api/excel/parse`

Dynamically parses any Excel file using column headers as property names.

**Features:**

- Handles any Excel structure
- Uses column headers as keys in the resulting dictionary
- Skips empty rows and columns
- Preserves original data types (string, number, date, boolean)

**Response Example:**

```json
{
  "totalRows": 2,
  "data": [
    {
      "Name": "John Doe",
      "Email": "john@example.com",
      "Age": 30
    },
    {
      "Name": "Jane Smith", 
      "Email": "jane@example.com",
      "Age": 25
    }
  ]
}
```

### 2. Attribute-Based Mapping

**POST** `/api/excel/persons`

Maps Excel data to `PersonDto` objects using attribute configuration.

**Features:**

- Strongly-typed result
- Configurable column mapping
- Data validation and filtering
- Required field enforcement

**Response Example:**

```json
{
  "totalRows": 2,
  "filteredFrom": 3,
  "data": [
    {
      "name": "John Doe",
      "email": "john@example.com", 
      "age": 30
    },
    {
      "name": "Jane Smith",
      "email": "jane@example.com",
      "age": 25
    }
  ]
}
```

## 🎯 Creating Custom DTOs

### Step 1: Define Your DTO

Create a new DTO class with `ExcelColumnAttribute` decorations:

```csharp
using ExcelParser.Attributes;

public class EmployeeDto
{
    [ExcelColumn("Full Name", IsRequired = true, FilterNullOrEmpty = true)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Employee ID", IsRequired = true)]
    public int EmployeeId { get; set; }

    [ExcelColumn("Department", FilterNullOrEmpty = true)]
    public string Department { get; set; } = string.Empty;

    [ExcelColumn("Salary", FilterNullOrEmpty = false)]
    public decimal? Salary { get; set; }

    [ExcelColumn("Start Date")]
    public DateTime? StartDate { get; set; }
}
```

### Step 2: Create Controller Action

Add a new endpoint in `ExcelParseController`:

```csharp
[HttpPost("employees")]
public async Task<IActionResult> ParseAsEmployees(IFormFile? file)
{
    // ... validation logic ...
    
    var dynamicResults = await excelService.ParseAsync(stream);
    var mappedResults = mappingService.MapToDto<EmployeeDto>(dynamicResults);

    return Ok(new
    {
        TotalRows = mappedResults.Count,
        FilteredFrom = dynamicResults.Count,
        Data = mappedResults
    });
}
```

## ⚙️ ExcelColumnAttribute Configuration

| Property            | Type   | Default | Description                                           |
|---------------------|--------|---------|-------------------------------------------------------|
| `ColumnName`        | string | -       | Excel column header name (case-insensitive)           |
| `FilterNullOrEmpty` | bool   | `true`  | Skip records with null/empty values for this property |
| `IsRequired`        | bool   | `false` | Skip entire record if this property is null/empty     |

### Configuration Examples

```csharp
// Basic mapping
[ExcelColumn("Name")]
public string Name { get; set; }

// Required field - skip entire record if missing
[ExcelColumn("Employee ID", IsRequired = true)]
public int EmployeeId { get; set; }

// Allow null/empty values
[ExcelColumn("Middle Name", FilterNullOrEmpty = false)]
public string? MiddleName { get; set; }

// Complex configuration
[ExcelColumn("Annual Salary", IsRequired = true, FilterNullOrEmpty = true)]
public decimal Salary { get; set; }
```

## 🔧 Project Structure

```
ExcelParser/
├── Attributes/
│   └── ExcelColumnAttribute.cs      # Custom attribute for column mapping
├── Controllers/
│   └── ExcelParseController.cs      # API endpoints
├── Dtos/
│   └── PersonDto.cs                 # Example DTO implementation
├── Models/
│   └── DynamicExcelData.cs          # Dynamic data container
├── Services/
│   ├── ExcelService.cs              # Core Excel parsing logic
│   ├── AttributeMappingService.cs   # DTO mapping service
│   └── Interfaces/
│       ├── IExcelService.cs         # Excel service contract
│       └── IMappingService.cs       # Mapping service contract
├── Program.cs                       # Application entry point
└── ExcelParser.csproj              # Project configuration
```

## 📦 Dependencies

### Main Project

- **ExcelDataReader** (3.7.0): Core Excel file reading functionality
- **Microsoft.AspNetCore.OpenApi** (9.0.6): OpenAPI specification support
- **Swashbuckle.AspNetCore** (9.0.3): Swagger UI and documentation

### Test Project

- **NUnit** (4.2.2): Testing framework
- **NSubstitute** (5.1.0): Mocking library for creating test doubles
- **ClosedXML** (0.105.0): Excel file creation for test data generation

## 🔍 Supported File Types

- **Excel 97-2003** (.xls): `application/vnd.ms-excel`
- **Excel 2007+** (.xlsx): `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

## 🚨 Error Handling

The application includes comprehensive error handling:

- **File validation**: Checks for required files and valid content types
- **Row-level resilience**: Continues processing even if individual rows fail
- **Type conversion**: Graceful handling of type conversion errors
- **Logging**: Detailed logging for debugging and monitoring

## 🧪 Testing

### Comprehensive Unit Test Suite

The project includes a complete NUnit test suite with over 30 test cases using NSubstitute for mocking, covering:

#### ExcelServiceTests

- Valid Excel data parsing with proper .xlsx files (using ClosedXML)
- Type preservation across different data types (string, double, DateTime, bool)
- Empty row handling and null value processing
- Mixed data type support and automatic type conversion
- Stream position management and reset functionality
- Logging verification and information capture

#### AttributeMappingServiceTests

- DTO mapping with attribute configuration
- Case-insensitive column name matching
- Data filtering (`FilterNullOrEmpty` behavior)
- Required field validation (`IsRequired` enforcement)
- Type conversion and error handling
- Missing column scenarios

#### ExcelColumnAttributeTests

- Attribute property initialization and configuration
- Default value behavior
- AttributeUsage validation

#### ExcelParseControllerTests

- File validation (null, empty, invalid content types)
- API endpoint integration testing
- Response structure validation
- Service interaction verification

#### DynamicExcelDataTests

- Data storage and retrieval functionality
- Type preservation across different data types
- Key overwriting behavior and data integrity
- Null value handling and edge cases

### Test Data Generation

The test suite uses **ClosedXML** to create actual Excel files for testing rather than mocked data:

- **Real Excel Format**: Tests use proper .xlsx files with correct OOXML structure
- **Type Accuracy**: Ensures Excel's native data types (string, double, DateTime, bool) are tested
- **Header Validation**: Tests proper Excel header row processing
- **Data Integrity**: Verifies actual Excel reading capabilities match expected behavior

### Running Tests

```bash
# Run all tests
dotnet test

# Run specific test class
dotnet test --filter "ClassName=ExcelServiceTests"
```

### Manual Testing with Sample Files

The project includes sample Excel files:

- `persons.xlsx`: Valid test data
- `persons_wrong.xlsx`: Test data with validation issues

### Using Swagger UI

1. Run the application in Development mode
2. Navigate to `https://localhost:5000/swagger`
3. Use the interactive interface to test endpoints

## 🔮 Extensibility

The architecture supports easy extension:

1. **New DTOs**: Simply create classes with `ExcelColumnAttribute` decorations
2. **Custom Validation**: Extend `ExcelColumnAttribute` with additional validation properties
3. **Different File Formats**: Implement `IExcelService` for other file types
4. **Custom Mapping Logic**: Implement `IMappingService` for specialized mapping requirements

## 🤝 Contributing

1. Follow the established patterns for new DTOs and endpoints
2. Add appropriate logging for new functionality
3. Update this README for significant changes
4. Test with both sample files and edge cases 