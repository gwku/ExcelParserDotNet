using System.Reflection;
using ExcelParser.Attributes;
using ExcelParser.Models;
using ExcelParser.Services.Interfaces;

namespace ExcelParser.Services;

public class AttributeMappingService(ILogger<AttributeMappingService> logger) : IMappingService
{
    public List<T> MapToDto<T>(List<DynamicExcelData> dynamicData) where T : class, new()
    {
        var results = new List<T>();
        var properties = typeof(T).GetProperties()
            .Where(p => p.GetCustomAttribute<ExcelColumnAttribute>() != null)
            .ToList();

        if (properties.Count is 0)
        {
            logger.LogWarning("No properties with ExcelColumn attributes found in {ModelType}", typeof(T).Name);
            return results;
        }

        logger.LogInformation(
            "Mapping {RecordCount} records to {ModelType} using {PropertyCount} attributed properties",
            dynamicData.Count, typeof(T).Name, properties.Count);

        foreach (var data in dynamicData)
        {
            var instance = new T();
            var hasData = false;
            var shouldFilterRecord = false;

            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<ExcelColumnAttribute>()!;
                var key = data.Data.Keys.FirstOrDefault(k =>
                    string.Equals(k, attribute.ColumnName, StringComparison.OrdinalIgnoreCase));

                if (key is null) continue;

                var value = data.Data[key];

                if (attribute.FilterNullOrEmpty && IsNullOrEmpty(value))
                {
                    shouldFilterRecord = true;
                    break;
                }

                if (TrySetValue(property, instance, value))
                    hasData = true;
            }

            if (!shouldFilterRecord && hasData && HasRequiredFields(instance, properties))
                results.Add(instance);
        }

        logger.LogInformation("Successfully mapped {MappedCount} out of {TotalCount} records to {ModelType}",
            results.Count, dynamicData.Count, typeof(T).Name);

        return results;
    }

    private static bool IsNullOrEmpty(object? value)
    {
        return value == null || (value is string str && string.IsNullOrWhiteSpace(str));
    }

    private bool TrySetValue(PropertyInfo property, object instance, object? value)
    {
        if (value == null) return false;

        try
        {
            var convertedValue = Convert.ChangeType(value, property.PropertyType);
            property.SetValue(instance, convertedValue);
            return true;
        }
        catch (Exception ex)
        {
            logger.LogDebug("Failed to convert value '{Value}' to {PropertyType} for property {PropertyName}: {Error}",
                value, property.PropertyType.Name, property.Name, ex.Message);
            return false;
        }
    }

    private static bool HasRequiredFields(object instance, List<PropertyInfo> properties)
    {
        return properties
            .Where(p => p.GetCustomAttribute<ExcelColumnAttribute>()!.IsRequired)
            .All(p =>
            {
                var value = p.GetValue(instance);
                return !IsNullOrEmpty(value);
            });
    }
}