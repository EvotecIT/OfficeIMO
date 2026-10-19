using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ChartFormula = DocumentFormat.OpenXml.Drawing.Charts.Formula;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel.Utilities;

/// <summary>Compares projections using stable sheet names without changing the live package.</summary>
internal static class ExcelPackageReferenceNormalizer {
    internal static OpenXmlElement Copy(OpenXmlElement root, IReadOnlyDictionary<string, string>? names) {
        var clone = root.CloneNode(true);
        if (names == null || names.Count == 0) return clone;
        foreach (var leaf in clone.Descendants().Prepend(clone).OfType<OpenXmlLeafTextElement>()) {
            if (leaf is CellFormula or Formula or Formula1 or Formula2 or DefinedName or ChartFormula or OfficeFormula or CalculatedColumnFormula or TotalsRowFormula)
                leaf.Text = ExcelDocument.ReplaceSheetNameReferences(leaf.Text, names);
        }
        foreach (var link in clone.Descendants<Hyperlink>())
            if (link.Location?.Value is string location) link.Location = ExcelDocument.ReplaceSheetNameReferences(location, names);
        foreach (var source in clone.Descendants<WorksheetSource>())
            if (source.Sheet?.Value is string name && names.TryGetValue(name, out string? original)) source.Sheet = original;
        return clone;
    }
}
