namespace ExcelParser.Models;

public class DynamicExcelData
{
    public Dictionary<string, object> Data { get; } = [];

    public void SetValue(string key, object value)
    {
        Data[key] = value;
    }
}