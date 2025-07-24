using System.Text;
using ExcelDataReader;
using ExcelParser.Models;
using ExcelParser.Services.Interfaces;

namespace ExcelParser.Services;

public class ExcelService(ILogger<ExcelService> logger) : IExcelService
{
    public Task<List<DynamicExcelData>> ParseAsync(Stream excelStream)
    {
        logger.LogInformation("Starting dynamic Excel parsing");

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ResetStreamIfNeeded(excelStream);

        using var reader = ExcelReaderFactory.CreateReader(excelStream);
        var results = new List<DynamicExcelData>();
        var headers = new Dictionary<int, string>();

        for (var row = 0; reader.Read(); row++)
        {
            if (row == 0)
            {
                ProcessHeaderRow(reader, headers);
                continue;
            }

            if (!TryParseRow(reader, headers, out var data))
            {
                logger.LogWarning("Skipped empty or unprocessable row at index {RowIndex}.", row);
                continue;
            }

            results.Add(data);
            logger.LogDebug("Parsed row {RowIndex}.", row);
        }

        logger.LogInformation(
            "Completed parsing. Total rows processed (incl. header): {ProcessedRowCount} of {RowCount}",
            results.Count + 1, reader.RowCount);

        return Task.FromResult(results);
    }

    private void ResetStreamIfNeeded(Stream stream)
    {
        if (stream.Position == 0) return;

        logger.LogDebug("Resetting stream position to 0.");
        stream.Position = 0;
    }

    private void ProcessHeaderRow(IExcelDataReader reader, Dictionary<int, string> headers)
    {
        logger.LogDebug("Processing header row.");

        for (var col = 0; col < reader.FieldCount; col++)
        {
            var rawHeader = reader.GetValue(col)?.ToString()?.Trim();

            if (string.IsNullOrWhiteSpace(rawHeader))
            {
                logger.LogWarning("Empty header found at column {ColumnIndex}. Skipping this column.", col);
                continue;
            }

            headers[col] = rawHeader;
            logger.LogDebug("Found column {ColumnIndex}: '{Header}'", col, rawHeader);
        }
    }

    private static bool TryParseRow(
        IExcelDataReader reader,
        Dictionary<int, string> headers,
        out DynamicExcelData data)
    {
        data = new DynamicExcelData();
        var populated = false;

        for (var col = 0; col < reader.FieldCount; col++)
        {
            if (!headers.TryGetValue(col, out var header)) continue;

            var cellValue = reader.GetValue(col);

            // Handle different data types appropriately
            object? processedValue = cellValue switch
            {
                null => null,
                string str when string.IsNullOrWhiteSpace(str) => null,
                double d => d,
                int i => i,
                DateTime dt => dt,
                bool b => b,
                _ => cellValue.ToString()
            };

            if (processedValue is null) continue;

            data.SetValue(header, processedValue);
            populated = true;
        }

        return populated;
    }
}