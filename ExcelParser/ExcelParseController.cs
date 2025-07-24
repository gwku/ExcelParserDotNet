using ExcelParser.Dtos;
using ExcelParser.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace ExcelParser;

[ApiController]
[Route("api/excel")]
public class ExcelParseController(IExcelService excelService, IMappingService mappingService) : Controller
{
    /// <summary>
    ///     Upload and parse an Excel file dynamically using column headers as property names
    /// </summary>
    /// <param name="file">Excel file to parse</param>
    /// <returns>Dynamic data structure based on Excel content</returns>
    [HttpPost("parse")]
    [ProducesResponseType(typeof(List<Dictionary<string, object>>), StatusCodes.Status200OK)]
    public async Task<IActionResult> ParseExcel(IFormFile? file)
    {
        if (file is null || file.Length == 0)
            return BadRequest("File is required.");

        var permittedContentTypes = new[]
        {
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };

        if (!permittedContentTypes.Contains(file.ContentType))
            return BadRequest("Invalid content type. Please upload a valid Excel file.");

        using var stream = new MemoryStream();
        await file.CopyToAsync(stream);

        var results = await excelService.ParseAsync(stream);
        var response = results.Select(item => item.Data).ToList();

        return Ok(new
        {
            TotalRows = response.Count,
            Data = response
        });
    }

    /// <summary>
    ///     Upload and parse an Excel file, mapping it to Person DTOs using ExcelColumn attributes with filtering
    /// </summary>
    /// <param name="file">Excel file to parse</param>
    /// <returns>List of PersonDto objects with filtered data based on attribute settings</returns>
    [HttpPost("persons")]
    [ProducesResponseType(typeof(List<PersonDto>), StatusCodes.Status200OK)]
    public async Task<IActionResult> ParseAsPersons(IFormFile? file)
    {
        if (file is null || file.Length == 0)
            return BadRequest("File is required.");

        var permittedContentTypes = new[]
        {
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };

        if (!permittedContentTypes.Contains(file.ContentType))
            return BadRequest("Invalid content type. Please upload a valid Excel file.");

        using var stream = new MemoryStream();
        await file.CopyToAsync(stream);

        var dynamicResults = await excelService.ParseAsync(stream);
        var mappedResults = mappingService.MapToDto<PersonDto>(dynamicResults);

        return Ok(new
        {
            TotalRows = mappedResults.Count,
            FilteredFrom = dynamicResults.Count,
            Data = mappedResults
        });
    }
}