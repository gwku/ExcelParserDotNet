using ExcelParser.Attributes;

namespace ExcelParser.Dtos;

public class PersonDto
{
    [ExcelColumn("Name", IsRequired = true, FilterNullOrEmpty = true)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Email", FilterNullOrEmpty = true)]
    public string Email { get; set; } = string.Empty;

    [ExcelColumn("Age", FilterNullOrEmpty = true)]
    public int Age { get; set; }
}