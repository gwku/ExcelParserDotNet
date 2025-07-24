using ExcelParser.Attributes;

namespace ExcelParserTests;

[TestFixture]
public class ExcelColumnAttributeTests
{
    [Test]
    public void Constructor_WithColumnName_SetsColumnNameCorrectly()
    {
        // Arrange & Act
        var attribute = new ExcelColumnAttribute("TestColumn");

        // Assert
        Assert.That(attribute.ColumnName, Is.EqualTo("TestColumn"));
    }

    [Test]
    public void DefaultValues_AreSetCorrectly()
    {
        // Arrange & Act
        var attribute = new ExcelColumnAttribute("TestColumn");

        Assert.Multiple(() =>
        {
            // Assert
            Assert.That(attribute.FilterNullOrEmpty, Is.True); // Default value
            Assert.That(attribute.IsRequired, Is.False); // Default value
        });
    }

    [Test]
    public void FilterNullOrEmpty_CanBeSetToFalse()
    {
        // Arrange & Act
        var attribute = new ExcelColumnAttribute("TestColumn")
        {
            FilterNullOrEmpty = false
        };

        // Assert
        Assert.That(attribute.FilterNullOrEmpty, Is.False);
    }

    [Test]
    public void IsRequired_CanBeSetToTrue()
    {
        // Arrange & Act
        var attribute = new ExcelColumnAttribute("TestColumn")
        {
            IsRequired = true
        };

        // Assert
        Assert.That(attribute.IsRequired, Is.True);
    }

    [Test]
    public void AllProperties_CanBeSetTogether()
    {
        // Arrange & Act
        var attribute = new ExcelColumnAttribute("TestColumn")
        {
            FilterNullOrEmpty = false,
            IsRequired = true
        };

        Assert.Multiple(() =>
        {
            // Assert
            Assert.That(attribute.ColumnName, Is.EqualTo("TestColumn"));
            Assert.That(attribute.FilterNullOrEmpty, Is.False);
            Assert.That(attribute.IsRequired, Is.True);
        });
    }

    [Test]
    public void AttributeUsage_IsSetCorrectly()
    {
        // Arrange
        var attributeType = typeof(ExcelColumnAttribute);

        // Act
        var usageAttribute = (AttributeUsageAttribute)Attribute.GetCustomAttribute(
            attributeType, typeof(AttributeUsageAttribute));

        // Assert
        Assert.That(usageAttribute, Is.Not.Null);
        Assert.That(usageAttribute.ValidOn, Is.EqualTo(AttributeTargets.Property));
    }
}