using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OutlookUserPropertyTests {
    private static readonly byte[] MicrosoftTextFieldSample = new byte[] {
        0x03, 0x01, 0x01, 0x00, 0x00, 0x00, 0x45, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x0A, 0x00, 0x54, 0x00, 0x65, 0x00, 0x78, 0x00, 0x74, 0x00, 0x46, 0x00,
        0x69, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x64, 0x00, 0x31, 0x00, 0x0A, 0x54, 0x65, 0x78,
        0x74, 0x46, 0x69, 0x65, 0x6C, 0x64, 0x31, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x15, 0x00, 0x00, 0x00, 0x0A, 0x54, 0x00, 0x65, 0x00, 0x78, 0x00, 0x74, 0x00,
        0x46, 0x00, 0x69, 0x00, 0x65, 0x00, 0x6C, 0x00, 0x64, 0x00, 0x31, 0x00, 0x00, 0x00,
        0x00, 0x00
    };

    [Fact]
    public void ParsesMicrosoftPropertyDefinitionSample() {
        var document = new EmailDocument();
        document.Mapi.Set(MapiKnownProperties.PidLid.PropertyDefinitionStream, MicrosoftTextFieldSample);
        document.MapiProperties.Add(new MapiProperty(0x8001, MapiPropertyType.Unicode, "sample value",
            name: new MapiNamedProperty(MapiPropertySets.PublicStrings, "TextField1")));

        OutlookUserProperty property = Assert.Single(document.UserProperties);

        Assert.Equal(OutlookUserPropertyDefinitionState.Valid, document.UserProperties.DefinitionState);
        Assert.Equal("TextField1", property.Name);
        Assert.Equal(OutlookUserPropertyType.Text, property.FieldType);
        Assert.Equal((ushort)0x0008, property.Definition!.VariantType);
        Assert.Equal("sample value", property.Value);
        Assert.True(property.HasDefinition);
        Assert.True(property.HasValue);
    }

    [Fact]
    public void RoundTripsTypedUserPropertiesThroughMsg() {
        DateTimeOffset due = new DateTimeOffset(2026, 11, 4, 13, 45, 0, TimeSpan.Zero);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = "IPM.Note",
            Subject = "Custom fields",
            OutlookCodePage = 1250
        };
        source.UserProperties.Set("Customer name", "Żółw");
        source.UserProperties.Set("Approved", true);
        source.UserProperties.Set("Sequence", 42);
        source.UserProperties.Set("Score", 12.5d);
        source.UserProperties.Set("Cost", 19.95m);
        source.UserProperties.Set("Due", due);
        source.UserProperties.SetDuration("Effort", TimeSpan.FromMinutes(95));
        source.UserProperties.SetKeywords("Regions", new[] { "EMEA", "North America" });

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument result = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal(OutlookUserPropertyDefinitionState.Valid, result.UserProperties.DefinitionState);
        Assert.Equal(8, result.UserProperties.Count);
        Assert.Equal("Żółw", result.UserProperties.GetValueOrDefault<string>("customer NAME"));
        Assert.True(result.UserProperties.GetValueOrDefault<bool>("Approved"));
        Assert.Equal(42, result.UserProperties.GetValueOrDefault<int>("Sequence"));
        Assert.Equal(12.5d, result.UserProperties.GetValueOrDefault<double>("Score"));
        Assert.Equal(19.95m, result.UserProperties.GetValueOrDefault<decimal>("Cost"));
        Assert.Equal(due, result.UserProperties.GetValueOrDefault<DateTimeOffset>("Due"));
        Assert.Equal(TimeSpan.FromMinutes(95), result.UserProperties.GetValueOrDefault<TimeSpan>("Effort"));
        Assert.Equal(new[] { "EMEA", "North America" },
            result.UserProperties.GetValueOrDefault<string[]>("Regions"));
    }

    [Fact]
    public void AddingDefinitionPreservesExistingDefinitionBytes() {
        var document = new EmailDocument { OutlookCodePage = 1252 };
        document.Mapi.Set(MapiKnownProperties.PidLid.PropertyDefinitionStream, MicrosoftTextFieldSample);

        document.UserProperties.Set("Added", 7);

        byte[] updated = document.Mapi.GetValueOrDefault(MapiKnownProperties.PidLid.PropertyDefinitionStream)!;
        Assert.Equal(2U, BitConverter.ToUInt32(updated, 2));
        Assert.Equal(MicrosoftTextFieldSample.Skip(6), updated.Skip(6).Take(MicrosoftTextFieldSample.Length - 6));
        Assert.Equal(2, document.UserProperties.Definitions.Count);
    }

    [Fact]
    public void CorruptDefinitionStreamBlocksUnsafeRewriteTransactionally() {
        var document = new EmailDocument();
        byte[] corrupt = new byte[] { 0x03, 0x01, 0x01, 0x00, 0x00, 0x00, 0x45 };
        document.Mapi.Set(MapiKnownProperties.PidLid.PropertyDefinitionStream, corrupt);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            document.UserProperties.Set("Do not add", "value"));

        Assert.Contains("cannot be safely rewritten", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(corrupt,
            document.Mapi.GetValueOrDefault(MapiKnownProperties.PidLid.PropertyDefinitionStream));
        Assert.Null(document.MapiProperties.GetMapiProperty(MapiPropertySets.PublicStrings, "Do not add"));
        Assert.Equal(OutlookUserPropertyDefinitionState.Corrupt, document.UserProperties.DefinitionState);
    }

    [Fact]
    public void RemoveDeletesOnlyNamedValueAndDefinition() {
        var document = new EmailDocument();
        document.UserProperties.Set("Keep", "yes");
        document.UserProperties.Set("Remove", "no");

        Assert.True(document.UserProperties.Remove("remove"));

        OutlookUserProperty remaining = Assert.Single(document.UserProperties);
        Assert.Equal("Keep", remaining.Name);
        Assert.False(document.UserProperties.Remove("missing"));
    }

    [Fact]
    public void CategoriesOfferCaseInsensitiveSafeOperationsWhileRemainingAnIList() {
        var categories = new OutlookCategoryCollection { "Blue" };

        Assert.False(categories.AddIfMissing("blue"));
        Assert.True(categories.AddIfMissing("Green"));
        Assert.Equal(1, categories.RemoveAll("BLUE"));
        categories.ReplaceWith(new[] { "Red", "red", " Yellow " });

        Assert.Equal(new[] { "Red", "Yellow" }, categories);
    }
}
