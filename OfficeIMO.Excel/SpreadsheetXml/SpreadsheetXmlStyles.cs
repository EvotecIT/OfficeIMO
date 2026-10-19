using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

/// <summary>Maps number formats into the existing workbook stylesheet; other appearance is reported as loss.</summary>
internal static class SpreadsheetXmlStyles {
    internal static IReadOnlyDictionary<string, uint> Read(XElement? source, WorkbookPart workbook, SpreadsheetXmlContext context) {
        var result = new Dictionary<string, uint>(StringComparer.Ordinal);
        if (source == null) return result;
        context.Inspect(source, new[] { "Style" }, Array.Empty<string>());
        var definitions = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (XElement style in source.Elements(Ss + "Style")) {
            context.Options.CancellationToken.ThrowIfCancellationRequested();
            if (definitions.Count >= context.Options.MaxStyles) throw new InvalidDataException("XML Spreadsheet exceeds MaxStyles.");
            string name = (string?)style.Attribute(Ss + "ID") ?? throw new InvalidDataException("Style has no ID.");
            if (definitions.ContainsKey(name)) throw new InvalidDataException("Duplicate XML Spreadsheet style.");
            definitions.Add(name, style);
        }
        foreach (string name in definitions.Keys) Resolve(name, new HashSet<string>(StringComparer.Ordinal));
        return result;

        uint Resolve(string name, HashSet<string> active) {
            if (result.TryGetValue(name, out uint existing)) return existing;
            if (!active.Add(name) || active.Count > 128) throw new InvalidDataException("Cyclic or excessive XML Spreadsheet style inheritance.");
            if (!definitions.TryGetValue(name, out XElement? element)) throw new InvalidDataException("Unknown parent style.");
            context.Inspect(element, new[] { "NumberFormat" }, new[] { "ID", "Name", "Parent" });
            string? parent = (string?)element.Attribute(Ss + "Parent");
            uint inherited = parent != null ? Resolve(parent, active) : 0;
            XElement? format = element.Element(Ss + "NumberFormat");
            if (format != null) context.Inspect(format, Array.Empty<string>(), new[] { "Format" });
            string? code = (string?)format?.Attribute(Ss + "Format");
            uint value = code == null ? inherited : AddFormat(workbook, code);
            active.Remove(name); result[name] = value;
            return value;
        }
    }

    internal static uint EnsureDateStyle(WorkbookPart workbook, uint original) {
        Stylesheet styles = EnsureStyles(workbook);
        CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)original);
        if ((format.NumberFormatId?.Value ?? 0) != 0) return original;
        return AddFormat(workbook, "yyyy-mm-dd\\Thh:mm:ss");
    }

    private static uint AddFormat(WorkbookPart workbook, string code) {
        if (code.Length > 255) throw new InvalidDataException("XML Spreadsheet number format is too long.");
        Stylesheet styles = EnsureStyles(workbook);
        if (code.Length == 0 || code == "General") return 0;
        // XML Spreadsheet's named date format is an Excel display format, not an OOXML format code.
        code = code == "Short Date" ? "m/d/yy" : code == "Long Date" ? "dddd, mmmm dd, yyyy" : code == "General Date" ? "m/d/yy h:mm" : code;
        NumberingFormats formats = styles.NumberingFormats ??= new NumberingFormats();
        NumberingFormat? numberFormat = formats.Elements<NumberingFormat>().FirstOrDefault(item => item.FormatCode?.Value == code);
        if (numberFormat == null) {
            uint next = Math.Max(164, formats.Elements<NumberingFormat>().Select(item => item.NumberFormatId?.Value ?? 0).DefaultIfEmpty(163U).Max() + 1);
            numberFormat = new NumberingFormat { NumberFormatId = next, FormatCode = code };
            formats.Append(numberFormat); formats.Count = (uint)formats.ChildElements.Count;
        }
        var cells = styles.CellFormats!;
        var existing = cells.Elements<CellFormat>().Select((item, index) => new { item, index }).FirstOrDefault(item => item.item.NumberFormatId?.Value == numberFormat.NumberFormatId?.Value);
        if (existing != null) return (uint)existing.index;
        uint id = (uint)cells.ChildElements.Count;
        cells.Append(new CellFormat { NumberFormatId = numberFormat.NumberFormatId!.Value, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0, ApplyNumberFormat = true });
        cells.Count = (uint)cells.ChildElements.Count;
        return id;
    }

    private static Stylesheet EnsureStyles(WorkbookPart workbook) {
        WorkbookStylesPart part = workbook.WorkbookStylesPart ?? workbook.AddNewPart<WorkbookStylesPart>();
        part.Stylesheet ??= new Stylesheet();
        Stylesheet styles = part.Stylesheet;
        styles.Fonts ??= new Fonts(new Font()) { Count = 1 };
        styles.Fills ??= new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2 };
        styles.Borders ??= new Borders(new Border()) { Count = 1 };
        styles.CellStyleFormats ??= new CellStyleFormats(new CellFormat()) { Count = 1 };
        styles.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1 };
        return styles;
    }

    internal static IReadOnlyList<string> Export(Stylesheet? source, SpreadsheetXmlContext context) {
        if (source == null) return new[] { "General" };
        Stylesheet defaults = ExcelDocument.CreateDefaultStylesheet();
        if (!Equivalent(source.Fonts, defaults.Fonts) || !Equivalent(source.Fills, defaults.Fills)
            || !Equivalent(source.Borders, defaults.Borders) || !Equivalent(source.CellStyleFormats, defaults.CellStyleFormats))
            context.Omitted("font, fill, border or base-style definitions");
        foreach (var child in source.ChildElements) {
            if (child is not Fonts && child is not Fills && child is not Borders && child is not CellStyleFormats
                && child is not CellFormats && child is not NumberingFormats)
                context.Omitted("stylesheet " + child.LocalName);
        }
        var formats = source.NumberingFormats?.Elements<NumberingFormat>().ToDictionary(item => item.NumberFormatId!.Value, item => item.FormatCode!.Value!)
            ?? new Dictionary<uint, string>();
        var result = new List<string>();
        foreach (CellFormat cell in source.CellFormats?.Elements<CellFormat>() ?? Enumerable.Empty<CellFormat>()) {
            if (result.Count >= context.Options.MaxStyles) throw new InvalidDataException("Workbook exceeds MaxStyles.");
            uint number = cell.NumberFormatId?.Value ?? 0;
            string? code = formats.TryGetValue(number, out string? custom) ? custom : ExcelBuiltInNumberFormats.GetCode(number);
            if (code == null) { context.Omitted("unmapped built-in number format " + number); code = "General"; }
            if ((cell.FontId?.Value ?? 0) != 0 || (cell.FillId?.Value ?? 0) != 0 || (cell.BorderId?.Value ?? 0) != 0
                || (cell.FormatId?.Value ?? 0) != 0 || cell.HasChildren)
                context.Omitted("cell appearance other than number formats");
            result.Add(code);
        }
        if (result.Count == 0) result.Add("General");
        return result;
    }

    private static bool Equivalent(DocumentFormat.OpenXml.OpenXmlElement? first, DocumentFormat.OpenXml.OpenXmlElement? second) =>
        first == null ? second == null : second != null && XNode.DeepEquals(XElement.Parse(first.OuterXml), XElement.Parse(second.OuterXml));
}
