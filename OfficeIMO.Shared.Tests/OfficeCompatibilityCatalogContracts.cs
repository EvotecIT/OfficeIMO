using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Security;
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

    [Fact]
    public void ProtectedContentCatalogIsDeterministicAndKeepsNonCryptographicProtectionDistinct() {
        OfficeProtectionCapabilityCatalog catalog = OfficeProtectionCapabilityCatalog.Current;

        string first = catalog.ToJson();
        string second = catalog.ToJson();
        using JsonDocument parsed = JsonDocument.Parse(first);

        Assert.Equal(first, second);
        Assert.Equal(catalog.Capabilities.Count, parsed.RootElement.GetProperty("capabilities").GetArrayLength());
        Assert.Equal(OfficeProtectionKind.AccessDeterrence, catalog.Get("pst-password").Kind);
        Assert.Equal(OfficeProtectionKind.EditingRestriction, catalog.Get("rtf-editing-restrictions").Kind);
        Assert.Equal(OfficeProtectionCoverageState.Blocked, catalog.Get("onenote-encrypted-revision").Mutate);
        Assert.Equal(OfficeProtectionCoverageState.Supported, catalog.Get("odf-password").Create);
        Assert.Equal(OfficeProtectionCoverageState.NotApplicable, catalog.Get("epub-font-obfuscation").Mutate);
        Assert.Equal(OfficeProtectionCoverageState.NotSupported, catalog.Get("smime-signature-msg-tnef").Create);
        Assert.Contains("| Inspect | Open | Create |", catalog.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void ProtectedContentCatalogEscapesEveryJsonControlCharacter() {
        var row = new OfficeProtectionCapability(
            "control-row", "EML\tformat", "OfficeIMO.Email", OfficeProtectionKind.DigitalSignature,
            OfficeProtectionCoverageState.Supported, OfficeProtectionCoverageState.Supported,
            OfficeProtectionCoverageState.NotSupported, OfficeProtectionCoverageState.NotApplicable,
            OfficeProtectionCoverageState.Preserved, OfficeProtectionCoverageState.NotApplicable,
            "Verify\u0001Api", "line\bfeed\f");
        var catalog = new OfficeProtectionCapabilityCatalog("control\u0002catalog", 1, new[] { row });

        using JsonDocument parsed = JsonDocument.Parse(catalog.ToJson());

        Assert.Equal("EML\tformat", parsed.RootElement.GetProperty("capabilities")[0].GetProperty("formatId").GetString());
    }
}
