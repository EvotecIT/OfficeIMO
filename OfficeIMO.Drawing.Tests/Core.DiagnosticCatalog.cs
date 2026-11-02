using System.Reflection;
using System.Text.Json;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeDiagnosticCatalogTests {
    [Fact]
    public void MachineReadableMigrationMapNamesRealUniqueCatalogConstants() {
        using JsonDocument map = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "TestAssets", "diagnostic-code-map.json")));
        var fields = typeof(OfficeDiagnosticCodes).GetFields(BindingFlags.Public | BindingFlags.Static)
            .ToDictionary(field => field.Name, field => (string)field.GetRawConstantValue()!);
        var oldCodes = new HashSet<string>(StringComparer.Ordinal);
        var mappings = new Dictionary<string, string>(StringComparer.Ordinal);
        var mappedCodes = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement entry in map.RootElement.GetProperty("entries").EnumerateArray()) {
            string oldCode = entry.GetProperty("oldCode").GetString()!;
            string code = entry.GetProperty("catalogCode").GetString()!;
            string member = entry.GetProperty("member").GetString()!;
            Assert.True(oldCodes.Add(oldCode), "Duplicate old code: " + oldCode);
            Assert.Equal(code, fields[member]);
            mappedCodes.Add(code);
            mappings.Add(oldCode, code);
        }
        Assert.Equal(fields.Count, fields.Values.Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(fields.Count, mappedCodes.Count);
        Assert.Equal("chart-flattened", OfficeDiagnosticCodes.ChartFlattened);
        Assert.Equal(OfficeDiagnosticCodes.PdfExcelPageContentOmitted, mappings["PDF_EXCEL_PAGE_CONTENT_OMITTED"]);
        Assert.Equal(OfficeDiagnosticCodes.PdfExcelTablesTruncated, mappings["PDF_EXCEL_TABLES_TRUNCATED"]);
        Assert.Equal(OfficeDiagnosticCodes.PdfPowerpointTablesTruncated, mappings["PDF_POWERPOINT_TABLES_TRUNCATED"]);
        Assert.Equal(OfficeDiagnosticCodes.PdfPowerpointEditableContentOmitted, mappings["PDF_POWERPOINT_EDITABLE_CONTENT_OMITTED"]);
        Assert.Equal(OfficeDiagnosticCodes.DocOlePropertiesUnreadable, mappings["DOC-OLE-PROPERTIES-UNREADABLE"]);
        foreach (string helperCode in new[] { "table-content", "unsupported-effect", "unsupported-outline-dash",
            "invisible-outline", "occupied-field", "occluded-outline", "existing-widget", "existing-annotation",
            "low-confidence", "overlapping-proposals", "ambiguous-label" }) {
            Assert.Equal(helperCode, mappings[helperCode]);
        }
        Assert.Contains(map.RootElement.GetProperty("entries").EnumerateArray(),
            entry => entry.GetProperty("oldCode").GetString() == "chart-projected"
                && entry.GetProperty("catalogCode").GetString() == OfficeDiagnosticCodes.ChartFlattened);
    }

    [Fact]
    public void SharedChartVocabularyRetainsExistingRendererAndNativeAuthoringFamilies() {
        foreach (OfficeChartType legacy in Enum.GetValues(typeof(OfficeChartType))) {
            Assert.True(Enum.TryParse(legacy.ToString(), out OfficeChartType shared));
            Assert.Equal((int)legacy, (int)shared);
        }
        foreach (OfficeIMO.OfficeChartType native in Enum.GetValues(typeof(OfficeIMO.OfficeChartType))) {
            Assert.True(Enum.TryParse(native.ToString(), out OfficeChartType _));
        }
        foreach (OfficeIMO.Excel.ExcelModernChartType native in Enum.GetValues(typeof(OfficeIMO.Excel.ExcelModernChartType))) {
            if (native == OfficeIMO.Excel.ExcelModernChartType.Unsupported) continue;
            Assert.True(Enum.TryParse(native.ToString(), out OfficeChartType _));
        }
        Assert.Equal(OfficeChartType.ColumnClustered, OfficeChartType.Column);
        Assert.Equal(OfficeChartType.BarClustered, OfficeChartType.Bar);
    }
}
