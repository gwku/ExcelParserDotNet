namespace ExcelParser.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute(string columnName) : Attribute
{
    /// <summary>
    ///     The Excel column name to map from (case-insensitive)
    /// </summary>
    public string ColumnName { get; } = columnName;

    /// <summary>
    ///     Whether to filter out null/empty values for this property
    /// </summary>
    public bool FilterNullOrEmpty { get; set; } = true;

    /// <summary>
    ///     Whether this property is required - if missing or null, skip the entire record
    /// </summary>
    public bool IsRequired { get; set; }
}