using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class OfficeCompatibilityCatalogContractTests {
    [Fact]
    public void BinaryFormatCatalogsExposeUniqueStableRowsAndValidFormatReferences() {
        OfficeCapabilityCatalog[] catalogs = {
            WordCompatibilityCatalog.Current,
            ExcelCompatibilityCatalog.Xls,
            ExcelCompatibilityCatalog.Xlsb,
            PowerPointCompatibilityCatalog.Current
        };
        var knownFormats = WordFormatCatalog.All
            .Concat(ExcelFormatCatalog.All)
            .Concat(PowerPointFormatCatalog.All)
            .Select(format => format.Id)
            .ToHashSet(StringComparer.Ordinal);

        foreach (OfficeCapabilityCatalog catalog in catalogs) {
            Assert.NotEmpty(catalog.Capabilities);
            Assert.Equal(
                catalog.Capabilities.Count,
                catalog.Capabilities.Select(capability => capability.Id).Distinct(StringComparer.Ordinal).Count());
            Assert.All(catalog.Capabilities, capability => {
                Assert.Contains(capability.FormatId, knownFormats);
                if (capability.GetState(OfficeCapabilityLane.LegacyToModern) == OfficeCapabilityCoverageState.Dropped
                    || capability.GetState(OfficeCapabilityLane.ModernToLegacy) == OfficeCapabilityCoverageState.Dropped) {
                    Assert.NotEqual(OfficeCompatibilityImpact.None, capability.AffectedFidelity);
                    Assert.False(string.IsNullOrWhiteSpace(capability.Note));
                }
            });
        }
    }

    [Fact]
    public void CapabilitySerializationIsDeterministicAndMachineReadable() {
        OfficeCapabilityCatalog catalog = ExcelCompatibilityCatalog.Xlsb;

        string first = catalog.ToJson();
        string second = catalog.ToJson();

        Assert.Equal(first, second);
        using JsonDocument parsed = JsonDocument.Parse(first);
        Assert.Equal(catalog.Id, parsed.RootElement.GetProperty("id").GetString());
        Assert.Equal(catalog.SchemaVersion, parsed.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(catalog.Capabilities.Count, parsed.RootElement.GetProperty("capabilities").GetArrayLength());
        Assert.Contains("| Legacy import |", catalog.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointSharedCatalogRetainsStaticVisualAndOpaqueDistinctions() {
        OfficeCapability chart = PowerPointCompatibilityCatalog.Current.Get("PowerPoint.Ppt.Charts");
        OfficeCapability unknown = PowerPointCompatibilityCatalog.Current.Get("PowerPoint.Ppt.UnknownRecordsAndStreams");

        Assert.Equal(OfficeCapabilityCoverageState.Rasterized, chart.ModernToLegacy);
        Assert.True(chart.AffectedFidelity.HasFlag(OfficeCompatibilityImpact.Editability));
        Assert.Equal(OfficeCapabilityCoverageState.PreservedOpaque, unknown.LegacyRoundTrip);
        Assert.True(unknown.AffectedFidelity.HasFlag(OfficeCompatibilityImpact.Carrier));
    }
}
