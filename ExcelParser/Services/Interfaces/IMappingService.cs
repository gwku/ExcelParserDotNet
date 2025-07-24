using ExcelParser.Models;

namespace ExcelParser.Services.Interfaces;

public interface IMappingService
{
    /// <summary>
    ///     Maps dynamic Excel data to a DTO type using ExcelColumn attributes, with filtering based on attribute settings
    /// </summary>
    /// <typeparam name="T">The target DTO type with ExcelColumn attributes</typeparam>
    /// <param name="dynamicData">List of dynamic Excel data</param>
    /// <returns>List of mapped and filtered DTO instances</returns>
    List<T> MapToDto<T>(List<DynamicExcelData> dynamicData) where T : class, new();
}