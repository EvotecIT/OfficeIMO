using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

/// <summary>Maps worksheet views, page setup and ordinary protection flags without activating external content.</summary>
internal static partial class SpreadsheetXmlOptionsCodec {
    internal static void Read(XElement worksheetSource, ExcelSheet sheet, SpreadsheetXmlContext context) {
        XElement source = worksheetSource.Element(X + "WorksheetOptions") ?? new XElement(X + "WorksheetOptions");
        Inspect(source, context, "Visible", "PageSetup", "Print", "Selected", "ProtectObjects", "ProtectScenarios", "ProtectContents", "DoNotDisplayGridlines", "DoNotDisplayHeadings", "Zoom", "FitToPage", "FreezePanes", "FrozenNoSplit", "SplitHorizontal", "SplitVertical", "TopRowBottomPane", "LeftColumnRightPane", "ActivePane", "Panes", "EnableSelection", "AllowDeleteCols", "AllowDeleteRows", "AllowInsertCols", "AllowInsertRows", "AllowSizeCols", "AllowSizeRows", "AllowFormatCells", "AllowInsertHyperlinks", "AllowFilter", "AllowSort", "AllowUsePivotTables");
        var worksheet = sheet.WorksheetPart.Worksheet!;
        var views = worksheet.GetFirstChild<SheetViews>() ?? new SheetViews(new SheetView { WorkbookViewId = 0 });
        var view = views.GetFirstChild<SheetView>()!;
        if (source.Element(X + "DoNotDisplayGridlines") != null) view.ShowGridLines = false;
        if (source.Element(X + "DoNotDisplayHeadings") != null) view.ShowRowColHeaders = false;
        if (source.Element(X + "Selected") != null) view.TabSelected = true;
        if (source.Element(X + "Zoom") is XElement zoom) view.ZoomScale = Integer(zoom.Value, 10, 400);
        if (source.Element(X + "FreezePanes") != null) {
            double horizontal = Number(source.Element(X + "SplitHorizontal")?.Value ?? "0", 65536), vertical = Number(source.Element(X + "SplitVertical")?.Value ?? "0", 256);
            uint top = Integer(source.Element(X + "TopRowBottomPane")?.Value ?? horizontal.ToString(CultureInfo.InvariantCulture), 0, 65535);
            uint left = Integer(source.Element(X + "LeftColumnRightPane")?.Value ?? vertical.ToString(CultureInfo.InvariantCulture), 0, 255);
            view.Pane = new Pane { State = PaneStateValues.Frozen, HorizontalSplit = vertical, VerticalSplit = horizontal, TopLeftCell = A1.CellReference((int)top + 1, (int)left + 1), ActivePane = Pane(source.Element(X + "ActivePane")?.Value ?? (vertical > 0 && horizontal > 0 ? "0" : vertical > 0 ? "1" : "2")) };
            if (source.Element(X + "Panes") != null) context.Omitted("pane selections");
        } else if (source.Element(X + "SplitHorizontal") != null || source.Element(X + "SplitVertical") != null) context.Omitted("unfrozen split panes");
        if (worksheetSource.Attribute(Ss + "RightToLeft") is XAttribute rtl) view.RightToLeft = SpreadsheetXmlGeometry.Flag(rtl.Value);
        worksheet.AddChild(views, true);
        if (source.Element(X + "PageSetup") is XElement setup) ReadPageSetup(setup, worksheet, context);
        if (source.Element(X + "Print") is XElement print) ReadPrint(print, worksheet, context);
        if (source.Element(X + "FitToPage") != null) {
            var properties = worksheet.SheetProperties ??= new SheetProperties();
            properties.PageSetupProperties ??= new PageSetupProperties(); properties.PageSetupProperties.FitToPage = true;
        }
        ReadProtection(worksheetSource, source, worksheet);
        Utilities.ExcelXmlPartWriter.SavePreservingLineEndings(sheet.WorksheetPart, worksheet);
    }

    private static void ReadPageSetup(XElement source, Worksheet worksheet, SpreadsheetXmlContext context) {
        Inspect(source, context, "Layout", "Header", "Footer", "PageMargins");
        var margins = worksheet.GetFirstChild<PageMargins>() ?? new PageMargins { Left = .7, Right = .7, Top = .75, Bottom = .75, Header = .3, Footer = .3 };
        if (source.Element(X + "PageMargins") is XElement pageMargins) {
            Attributes(pageMargins, context, "Left", "Right", "Top", "Bottom");
            if (pageMargins.Attribute(X + "Left") is XAttribute left) margins.Left = Number(left.Value, 49);
            if (pageMargins.Attribute(X + "Right") is XAttribute right) margins.Right = Number(right.Value, 49);
            if (pageMargins.Attribute(X + "Top") is XAttribute top) margins.Top = Number(top.Value, 49);
            if (pageMargins.Attribute(X + "Bottom") is XAttribute bottom) margins.Bottom = Number(bottom.Value, 49);
        }
        var headerFooter = new HeaderFooter();
        foreach (string name in new[] { "Header", "Footer" }) if (source.Element(X + name) is XElement item) {
            Attributes(item, context, "Margin", "Data");
            if (item.Attribute(X + "Margin") is XAttribute margin) {
                if (name == "Header") margins.Header = Number(margin.Value, 49); else margins.Footer = Number(margin.Value, 49);
            }
            if (item.Attribute(X + "Data") is XAttribute text) {
                if (text.Value.Length > 255) throw new InvalidDataException("XML Spreadsheet header/footer exceeds 255 characters.");
                if (name == "Header") headerFooter.OddHeader = new OddHeader(text.Value); else headerFooter.OddFooter = new OddFooter(text.Value);
            }
        }
        worksheet.AddChild(margins, true);
        if (headerFooter.HasChildren) worksheet.AddChild(headerFooter, true);
        if (source.Element(X + "Layout") is XElement layout) {
            Attributes(layout, context, "Orientation", "CenterHorizontal", "CenterVertical", "StartPageNumber");
            var page = worksheet.GetFirstChild<PageSetup>() ?? new PageSetup();
            if ((string?)layout.Attribute(X + "Orientation") is string orientation) page.Orientation = orientation == "Landscape" ? OrientationValues.Landscape : orientation == "Portrait" ? OrientationValues.Portrait : throw new InvalidDataException("Invalid page orientation.");
            if (layout.Attribute(X + "StartPageNumber") is XAttribute first) { page.FirstPageNumber = Integer(first.Value, 1, 32767); page.UseFirstPageNumber = true; }
            worksheet.AddChild(page, true);
            var options = worksheet.GetFirstChild<PrintOptions>() ?? new PrintOptions();
            if (layout.Attribute(X + "CenterHorizontal") is XAttribute horizontal) options.HorizontalCentered = SpreadsheetXmlGeometry.Flag(horizontal.Value);
            if (layout.Attribute(X + "CenterVertical") is XAttribute vertical) options.VerticalCentered = SpreadsheetXmlGeometry.Flag(vertical.Value);
            worksheet.AddChild(options, true);
        }
    }

    private static void ReadPrint(XElement source, Worksheet worksheet, SpreadsheetXmlContext context) {
        Inspect(source, context, "ValidPrinterInfo", "PaperSizeIndex", "Scale", "FitWidth", "FitHeight", "HorizontalResolution", "VerticalResolution", "Gridlines", "RowColHeadings", "BlackAndWhite", "DraftQuality", "LeftToRight");
        var page = worksheet.GetFirstChild<PageSetup>() ?? new PageSetup();
        if (source.Element(X + "PaperSizeIndex") is XElement paper) page.PaperSize = Integer(paper.Value, 1, 118);
        if (source.Element(X + "Scale") is XElement scale) page.Scale = Integer(scale.Value, 10, 400);
        if (source.Element(X + "FitWidth") is XElement width) page.FitToWidth = Integer(width.Value, 0, 32767);
        if (source.Element(X + "FitHeight") is XElement height) page.FitToHeight = Integer(height.Value, 0, 32767);
        if (source.Element(X + "HorizontalResolution") is XElement horizontal) page.HorizontalDpi = Integer(horizontal.Value, 1, 65535);
        if (source.Element(X + "VerticalResolution") is XElement vertical) page.VerticalDpi = Integer(vertical.Value, 1, 65535);
        if (source.Element(X + "BlackAndWhite") != null) page.BlackAndWhite = true;
        if (source.Element(X + "DraftQuality") != null) page.Draft = true;
        if (source.Element(X + "LeftToRight") != null) page.PageOrder = PageOrderValues.OverThenDown;
        worksheet.AddChild(page, true);
        var options = worksheet.GetFirstChild<PrintOptions>() ?? new PrintOptions();
        if (source.Element(X + "Gridlines") != null) options.GridLines = true;
        if (source.Element(X + "RowColHeadings") != null) options.Headings = true;
        worksheet.AddChild(options, true);
    }

    internal static void Write(XmlWriter writer, Sheet sheet, WorksheetPart? part, SpreadsheetXmlContext context) {
        var result = new XElement(X + "WorksheetOptions");
        if (sheet.State?.Value == SheetStateValues.Hidden || sheet.State?.Value == SheetStateValues.VeryHidden) result.Add(new XElement(X + "Visible", sheet.State.Value == SheetStateValues.Hidden ? "SheetHidden" : "SheetVeryHidden"));
        Worksheet? worksheet = part?.Worksheet;
        var view = worksheet?.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
        if (view?.ShowGridLines?.Value == false) result.Add(new XElement(X + "DoNotDisplayGridlines"));
        if (view?.ShowRowColHeaders?.Value == false) result.Add(new XElement(X + "DoNotDisplayHeadings"));
        if (view?.TabSelected?.Value == true) result.Add(new XElement(X + "Selected"));
        if (view?.ZoomScale != null) result.Add(new XElement(X + "Zoom", view.ZoomScale.Value));
        if (view?.Pane is DocumentFormat.OpenXml.Spreadsheet.Pane pane) {
            if (pane.State?.Value == PaneStateValues.Frozen || pane.State?.Value == PaneStateValues.FrozenSplit) {
                result.Add(new XElement(X + "FreezePanes"), new XElement(X + "FrozenNoSplit"), new XElement(X + "SplitHorizontal", pane.VerticalSplit?.Value ?? 0), new XElement(X + "SplitVertical", pane.HorizontalSplit?.Value ?? 0));
                if (ExcelReference.TryParse(pane.TopLeftCell?.Value ?? "A1", out var address) && address != null) result.Add(new XElement(X + "TopRowBottomPane", address.Start.Row - 1), new XElement(X + "LeftColumnRightPane", address.Start.Column - 1));
                result.Add(new XElement(X + "ActivePane", pane.ActivePane?.InnerText switch { "topRight" => "1", "bottomLeft" => "2", "topLeft" => "3", _ => "0" }));
            } else context.Omitted("unfrozen split panes");
        }
        WritePrint(result, worksheet, context);
        if (worksheet?.GetFirstChild<SheetProtection>() is SheetProtection protection) {
            WriteProtection(result, protection, context);
        }
        if (result.HasElements) result.WriteTo(writer);
    }

    private static void WritePrint(XElement target, Worksheet? worksheet, SpreadsheetXmlContext context) {
        var setup = new XElement(X + "PageSetup");
        var margins = worksheet?.GetFirstChild<PageMargins>(); var page = worksheet?.GetFirstChild<PageSetup>(); var options = worksheet?.GetFirstChild<PrintOptions>();
        var headerFooter = worksheet?.GetFirstChild<HeaderFooter>();
        if (page != null || options != null) {
            var layout = new XElement(X + "Layout");
            if (page?.Orientation != null) layout.Add(new XAttribute(X + "Orientation", page.Orientation.Value == OrientationValues.Landscape ? "Landscape" : "Portrait"));
            if (page?.UseFirstPageNumber?.Value == true && page.FirstPageNumber != null) layout.Add(new XAttribute(X + "StartPageNumber", page.FirstPageNumber.Value));
            if (options?.HorizontalCentered != null) layout.Add(new XAttribute(X + "CenterHorizontal", options.HorizontalCentered.Value ? "1" : "0"));
            if (options?.VerticalCentered != null) layout.Add(new XAttribute(X + "CenterVertical", options.VerticalCentered.Value ? "1" : "0"));
            if (layout.HasAttributes) setup.Add(layout);
        }
        if (margins != null) setup.Add(new XElement(X + "PageMargins", new XAttribute(X + "Left", margins.Left?.Value ?? .7), new XAttribute(X + "Right", margins.Right?.Value ?? .7), new XAttribute(X + "Top", margins.Top?.Value ?? .75), new XAttribute(X + "Bottom", margins.Bottom?.Value ?? .75)));
        if (margins?.Header != null || headerFooter?.OddHeader != null) setup.Add(new XElement(X + "Header", new XAttribute(X + "Margin", margins?.Header?.Value ?? .3), new XAttribute(X + "Data", headerFooter?.OddHeader?.Text ?? "")));
        if (margins?.Footer != null || headerFooter?.OddFooter != null) setup.Add(new XElement(X + "Footer", new XAttribute(X + "Margin", margins?.Footer?.Value ?? .3), new XAttribute(X + "Data", headerFooter?.OddFooter?.Text ?? "")));
        if (headerFooter?.DifferentFirst?.Value == true || headerFooter?.DifferentOddEven?.Value == true) context.Omitted("first-page or even-page headers and footers");
        if (setup.HasElements) target.Add(setup);
        var print = new XElement(X + "Print");
        void Add(string name, object? value) { if (value != null) print.Add(new XElement(X + name, value)); }
        Add("PaperSizeIndex", page?.PaperSize?.Value); Add("Scale", page?.Scale?.Value); Add("FitWidth", page?.FitToWidth?.Value); Add("FitHeight", page?.FitToHeight?.Value);
        Add("HorizontalResolution", page?.HorizontalDpi?.Value); Add("VerticalResolution", page?.VerticalDpi?.Value);
        if (page?.BlackAndWhite?.Value == true) Add("BlackAndWhite", ""); if (page?.Draft?.Value == true) Add("DraftQuality", "");
        if (page?.PageOrder?.Value == PageOrderValues.OverThenDown) Add("LeftToRight", "");
        if (options?.GridLines?.Value == true) Add("Gridlines", ""); if (options?.Headings?.Value == true) Add("RowColHeadings", "");
        if (print.HasElements) target.Add(print);
        if (worksheet?.SheetProperties?.PageSetupProperties?.FitToPage?.Value == true) target.Add(new XElement(X + "FitToPage"));
    }

    private static PaneValues Pane(string value) => value switch { "0" => PaneValues.BottomRight, "1" => PaneValues.TopRight, "2" => PaneValues.BottomLeft, "3" => PaneValues.TopLeft, _ => throw new InvalidDataException("Invalid pane index.") };
    private static double Number(string value, double maximum) => SpreadsheetXmlGeometry.Dimension(value, maximum);
    private static uint Integer(string value, uint minimum, uint maximum) => uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out uint parsed) && parsed >= minimum && parsed <= maximum ? parsed : throw new InvalidDataException("Invalid XML Spreadsheet option.");
    private static void Inspect(XElement element, SpreadsheetXmlContext context, params string[] supported) => context.Inspect(element, supported.Select(name => (X + name).ToString()).ToArray(), Array.Empty<string>());
    private static void Attributes(XElement element, SpreadsheetXmlContext context, params string[] supported) {
        foreach (var attribute in element.Attributes()) if (!attribute.IsNamespaceDeclaration && (attribute.Name.Namespace != X || !supported.Contains(attribute.Name.LocalName))) context.Omitted("page option attribute " + attribute.Name);
        foreach (var child in element.Elements()) context.Omitted("page option " + child.Name);
    }
}
