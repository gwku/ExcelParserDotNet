using ExcelParser.Models;

namespace ExcelParser.Services.Interfaces;

public interface IExcelService
{
    /// <summary>
    ///     Dynamically parses an Excel file using the column headers as property names
    /// </summary>
    /// <param name="excelStream">The Excel file stream</param>
    /// <returns>List of dynamic data objects containing all Excel data</returns>
    Task<List<DynamicExcelData>> ParseAsync(Stream excelStream);
}