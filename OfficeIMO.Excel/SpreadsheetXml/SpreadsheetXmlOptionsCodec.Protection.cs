using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

internal static partial class SpreadsheetXmlOptionsCodec {
    private static readonly IReadOnlyDictionary<string, string> ProtectionPermissions = new Dictionary<string, string> {
        ["AllowDeleteCols"] = "deleteColumns", ["AllowDeleteRows"] = "deleteRows", ["AllowInsertCols"] = "insertColumns",
        ["AllowInsertRows"] = "insertRows", ["AllowSizeCols"] = "formatColumns", ["AllowSizeRows"] = "formatRows",
        ["AllowFormatCells"] = "formatCells", ["AllowInsertHyperlinks"] = "insertHyperlinks", ["AllowFilter"] = "autoFilter",
        ["AllowSort"] = "sort", ["AllowUsePivotTables"] = "pivotTables"
    };

    private static void ReadProtection(XElement worksheetSource, XElement source, Worksheet worksheet) {
        if (worksheetSource.Attribute(Ss + "Protected") is not XAttribute protect || !SpreadsheetXmlGeometry.Flag(protect.Value)) return;
        var value = new SheetProtection { Sheet = true,
            Objects = source.Element(X + "ProtectObjects") is XElement objects && SpreadsheetXmlGeometry.Flag(objects.Value),
            Scenarios = source.Element(X + "ProtectScenarios") is XElement scenarios && SpreadsheetXmlGeometry.Flag(scenarios.Value) };
        foreach (var permission in ProtectionPermissions)
            value.SetAttribute(new OpenXmlAttribute(permission.Value, "", source.Element(X + permission.Key) != null ? "0" : "1"));
        string? selection = source.Element(X + "EnableSelection")?.Value;
        if (selection is not null and not "UnlockedCells" and not "NoSelection") throw new InvalidDataException("Invalid worksheet selection permission.");
        value.SelectLockedCells = selection is "UnlockedCells" or "NoSelection";
        value.SelectUnlockedCells = selection == "NoSelection";
        worksheet.AddChild(value, true);
    }

    private static void WriteProtection(XElement target, SheetProtection protection, SpreadsheetXmlContext context) {
        target.Add(new XElement(X + "ProtectObjects", protection.Objects?.Value == true ? "True" : "False"),
            new XElement(X + "ProtectScenarios", protection.Scenarios?.Value == true ? "True" : "False"));
        foreach (var permission in ProtectionPermissions) {
            string? text = protection.GetAttribute(permission.Value, "").Value;
            if (text is "0" or "false") target.Add(new XElement(X + permission.Key));
        }
        if (protection.SelectLockedCells?.Value == true) {
            target.Add(new XElement(X + "EnableSelection", protection.SelectUnlockedCells?.Value == true ? "NoSelection" : "UnlockedCells"));
        } else if (protection.SelectUnlockedCells?.Value == true) context.Omitted("selecting only locked cells");
        InspectOpenXml(protection, context, new[] { "sheet", "objects", "scenarios", "selectLockedCells", "selectUnlockedCells" }.Concat(ProtectionPermissions.Values).ToArray());
    }

    private static void InspectOpenXml(OpenXmlElement element, SpreadsheetXmlContext context, params string[] attributes) {
        foreach (OpenXmlAttribute attribute in element.GetAttributes())
            if (attribute.NamespaceUri.Length != 0 || !attributes.Contains(attribute.LocalName)) context.Omitted(element.LocalName + " attribute " + attribute.LocalName);
    }

    internal static void InspectNested(Worksheet worksheet, SpreadsheetXmlContext context) {
        if (worksheet.SheetProperties is SheetProperties properties) {
            InspectOpenXml(properties, context);
            foreach (var child in properties.ChildElements) {
                if (child is PageSetupProperties page) InspectOpenXml(page, context, "fitToPage");
                else context.Omitted("sheetPr/" + child.LocalName);
            }
        }
        if (worksheet.GetFirstChild<SheetFormatProperties>() is SheetFormatProperties format)
            InspectOpenXml(format, context, "baseColWidth", "defaultColWidth", "defaultRowHeight", "customHeight");
        if (worksheet.GetFirstChild<PageSetup>() is PageSetup setup)
            InspectOpenXml(setup, context, "paperSize", "orientation", "firstPageNumber", "useFirstPageNumber", "scale", "fitToWidth", "fitToHeight", "horizontalDpi", "verticalDpi", "blackAndWhite", "draft", "pageOrder");
        if (worksheet.GetFirstChild<PrintOptions>() is PrintOptions print) InspectOpenXml(print, context, "gridLines", "headings", "horizontalCentered", "verticalCentered");
        if (worksheet.GetFirstChild<PageMargins>() is PageMargins margins) InspectOpenXml(margins, context, "left", "right", "top", "bottom", "header", "footer");
        if (worksheet.GetFirstChild<HeaderFooter>() is HeaderFooter headerFooter) {
            InspectOpenXml(headerFooter, context);
            foreach (var child in headerFooter.ChildElements) if (child is not OddHeader && child is not OddFooter) context.Omitted("headerFooter/" + child.LocalName);
        }
        foreach (var view in worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>() ?? Enumerable.Empty<SheetView>())
            foreach (var child in view.ChildElements) if (child is not DocumentFormat.OpenXml.Spreadsheet.Pane) context.Omitted("sheetView/" + child.LocalName);
    }
}
