namespace ExcelParser.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    public ExcelColumnAttribute(string columnName)
    {
        ColumnName = columnName;
    }

    /// <summary>
    ///     The Excel column name to map from (case-insensitive)
    /// </summary>
    public string ColumnName { get; }

    /// <summary>
    ///     Whether to filter out null/empty values for this property
    /// </summary>
    public bool FilterNullOrEmpty { get; set; } = true;

    /// <summary>
    ///     Whether this property is required - if missing or null, skip the entire record
    /// </summary>
    public bool IsRequired { get; set; }
}