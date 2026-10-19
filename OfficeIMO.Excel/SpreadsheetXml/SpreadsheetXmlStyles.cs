using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.SpreadsheetXml.SpreadsheetXmlContext;

namespace OfficeIMO.Excel.SpreadsheetXml;

/// <summary>Maps XML Spreadsheet appearance into the canonical Open XML workbook stylesheet.</summary>
internal static partial class SpreadsheetXmlStyles {
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
            foreach (string child in new[] { "NumberFormat", "Font", "Interior", "Borders", "Alignment", "Protection" })
                if (element.Elements(Ss + child).Skip(1).Any()) throw new InvalidDataException("Duplicate style component.");
            context.Inspect(element, new[] { "NumberFormat", "Font", "Interior", "Borders", "Alignment", "Protection" }, new[] { "ID", "Name", "Parent" });
            string? parent = (string?)element.Attribute(Ss + "Parent");
            if (parent == null && name != "Default" && definitions.ContainsKey("Default")) parent = "Default";
            uint inherited = parent != null ? Resolve(parent, active) : 0;
            Stylesheet styles = EnsureStyles(workbook);
            var cell = (CellFormat)styles.CellFormats!.Elements<CellFormat>().ElementAt((int)inherited).CloneNode(true);
            ReadAppearance(element, cell, styles, context);
            XElement? format = element.Element(Ss + "NumberFormat");
            if (format != null) {
                context.Inspect(format, Array.Empty<string>(), new[] { "Format" });
                cell.NumberFormatId = NumberFormat(styles, (string?)format.Attribute(Ss + "Format") ?? "General");
                cell.ApplyNumberFormat = true;
            }
            if (name == "Default" && cell.FontId != null) {
                // Font zero defines worksheet column metrics in Open XML, independently of the cell XF.
                var font = (Font)styles.Fonts!.Elements<Font>().ElementAt((int)cell.FontId.Value).CloneNode(true);
                styles.Fonts.ReplaceChild(font, styles.Fonts.GetFirstChild<Font>()!);
            }
            uint value = AppendUnique(styles.CellFormats, cell);
            active.Remove(name); result[name] = value;
            return value;
        }
    }

    internal static uint EnsureDateStyle(WorkbookPart workbook, uint original, SpreadsheetXmlContext context) {
        if (context.DateStyles.TryGetValue(original, out uint existing)) return existing;
        Stylesheet styles = EnsureStyles(workbook);
        var format = (CellFormat)styles.CellFormats!.Elements<CellFormat>().ElementAt((int)original).CloneNode(true);
        if ((format.NumberFormatId?.Value ?? 0) != 0) return original;
        format.NumberFormatId = NumberFormat(styles, "yyyy-mm-dd\\Thh:mm:ss"); format.ApplyNumberFormat = true;
        return context.DateStyles[original] = AppendUnique(styles.CellFormats, format);
    }

    private static uint NumberFormat(Stylesheet styles, string code) {
        if (code.Length > 255) throw new InvalidDataException("XML Spreadsheet number format is too long.");
        if (code.Length == 0 || code == "General") return 0;
        code = code == "Short Date" ? "m/d/yy" : code == "Long Date" ? "dddd, mmmm dd, yyyy" : code == "General Date" ? "m/d/yy h:mm" : code;
        NumberingFormats formats = styles.NumberingFormats ??= new NumberingFormats();
        NumberingFormat? existing = formats.Elements<NumberingFormat>().FirstOrDefault(item => item.FormatCode?.Value == code);
        if (existing != null) return existing.NumberFormatId!.Value;
        uint next = Math.Max(164, formats.Elements<NumberingFormat>().Select(item => item.NumberFormatId?.Value ?? 0).DefaultIfEmpty(163U).Max() + 1);
        formats.Append(new NumberingFormat { NumberFormatId = next, FormatCode = code }); formats.Count = (uint)formats.ChildElements.Count;
        return next;
    }

    private static uint AppendUnique<T>(OpenXmlCompositeElement collection, T value) where T : OpenXmlElement {
        int index = 0;
        foreach (T item in collection.Elements<T>()) {
            if (item.OuterXml == value.OuterXml) return (uint)index;
            index++;
        }
        collection.Append(value);
        collection.SetAttribute(new OpenXmlAttribute("count", "", collection.ChildElements.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        return (uint)index;
    }

    private static Stylesheet EnsureStyles(WorkbookPart workbook) {
        WorkbookStylesPart part = workbook.WorkbookStylesPart ?? workbook.AddNewPart<WorkbookStylesPart>();
        part.Stylesheet ??= ExcelDocument.CreateDefaultStylesheet();
        return part.Stylesheet;
    }

    internal static string StyleId(uint index) => "s" + index;
}
