using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkLegacyXmlReader {
    private sealed class LegacyCellStyle {
        internal IWorkParagraphStyle? Paragraph { get; set; }
        internal IWorkNumberFormat? NumberFormat { get; set; }
        internal IWorkCellFill? Fill { get; set; }
        internal IWorkCellPadding? Padding { get; set; }
        internal IWorkCellVerticalAlignment? VerticalAlignment { get; set; }
    }

    /// <summary>Resolves selected and region styles once per source style, retaining sparse blank cells.</summary>
    private sealed class LegacyTableStyles {
        private readonly IWorkLegacyXmlCatalog _catalog;
        private readonly IWorkProjectionBudget _budget;
        private readonly Dictionary<XElement, LegacyCellStyle> _styles = new();
        private readonly LegacyCellStyle[] _regions;
        private static readonly LegacyCellStyle Empty = new();

        internal LegacyTableStyles(XElement model, IWorkLegacyXmlCatalog catalog, IWorkProjectionBudget budget) {
            _catalog = catalog;
            _budget = budget;
            XElement? tableStyle = catalog.Resolve(model.Element(Sf + "tabular-style")
                ?? model.Element(Sf + "tabular-style-ref"), "tabular-style");
            _regions = new[] {
                Region(tableStyle, "SFTDefaultBodyCellStyleProperty"),
                Region(tableStyle, "SFTDefaultHeaderRowCellStyleProperty"),
                Region(tableStyle, "SFTDefaultHeaderColumnCellStyleProperty"),
                Region(tableStyle, "SFTDefaultFooterRowCellStyleProperty")
            };
            TextStyles = new(_regions[0].Paragraph, _regions[1].Paragraph, _regions[2].Paragraph, _regions[3].Paragraph);
            FillStyles = new(_regions[0].Fill, _regions[1].Fill, _regions[2].Fill, _regions[3].Fill);
            Layouts = _regions.Select(style => (IWorkCellLayout?)new IWorkCellLayout(style.Padding,
                style.VerticalAlignment)).ToArray();
        }

        internal IWorkTableTextStyles TextStyles { get; }
        internal IWorkTableFillStyles FillStyles { get; }
        internal IReadOnlyList<IWorkCellLayout?> Layouts { get; }

        internal LegacyCellStyle ForCell(XElement record, int row, int column, int rows,
            int headerRows, int headerColumns, int footerRows) {
            string? identifier = (string?)record.Attribute(Sf + "s");
            if (identifier != null) return Read(_catalog.Style(record, identifier, "cell-style"));
            return _regions[row <= headerRows ? 1 : row > rows - footerRows ? 3 : column <= headerColumns ? 2 : 0];
        }

        private LegacyCellStyle Region(XElement? tableStyle, string property) =>
            Read(_catalog.Resolve(_catalog.Property(tableStyle, property)?.Elements().FirstOrDefault(), "cell-style"));

        private LegacyCellStyle Read(XElement? style) {
            if (style == null) return Empty;
            if (_styles.TryGetValue(style, out LegacyCellStyle? cached)) return cached;
            _budget.AddTableCatalogEntries(1);
            XElement? layout = _catalog.Resolve(_catalog.Property(style, "SFTCellStylePropertyLayoutStyle")
                ?.Elements().FirstOrDefault(), "layoutstyle");
            XElement? paragraph = _catalog.Resolve(_catalog.Property(style, "SFTCellStylePropertyParagraphStyle")
                ?.Elements().FirstOrDefault(), "paragraphstyle")
                ?? _catalog.Resolve(_catalog.Property(layout, "layoutParagraphStyle")?.Elements().FirstOrDefault(), "paragraphstyle");
            XElement? padding = _catalog.Resolve(_catalog.Property(layout, "padding")?.Elements().FirstOrDefault(), "padding");
            var result = new LegacyCellStyle {
                Paragraph = paragraph == null ? null : _catalog.ParagraphStyle(paragraph),
                NumberFormat = ReadLegacyNumberFormat(style, _catalog),
                Fill = ReadLegacyFill(_catalog.Property(style, "fill"), _catalog),
                Padding = ReadLegacyPadding(padding, _catalog),
                VerticalAlignment = ReadLegacyVertical(_catalog.Property(layout, "verticalAlignment"), _catalog)
            };
            _styles.Add(style, result);
            return result;
        }
    }

    private static IWorkCellPadding? ReadLegacyPadding(XElement? padding, IWorkLegacyXmlCatalog catalog) {
        if (padding == null) return null;
        double[] sides = new double[4];
        string[] names = { "left", "top", "right", "bottom" };
        for (int index = 0; index < names.Length; index++) {
            XAttribute? attribute = padding.Attribute(Sf + names[index]);
            if (attribute == null) continue;
            if (!TryLegacyNumber(attribute.Value, out sides[index]) || sides[index] < 0) {
                catalog.AddLoss("IWORK_LEGACY_TABLE_LAYOUT_OMITTED", "An unsupported Numbers '09 cell padding was omitted.");
                return null;
            }
        }
        return new IWorkCellPadding(sides[0], sides[1], sides[2], sides[3]);
    }

    private static IWorkCellVerticalAlignment? ReadLegacyVertical(XElement? property, IWorkLegacyXmlCatalog catalog) {
        if (property == null) return null;
        string? value = (string?)property.Element(Sf + "number")?.Attribute(Sfa + "number");
        IWorkCellVerticalAlignment? alignment = value switch {
            "0" => IWorkCellVerticalAlignment.Top,
            "1" => IWorkCellVerticalAlignment.Middle,
            "2" => IWorkCellVerticalAlignment.Bottom,
            _ => null
        };
        if (alignment == null) catalog.AddLoss("IWORK_LEGACY_TABLE_LAYOUT_OMITTED",
            "An unsupported Numbers '09 vertical cell alignment was omitted.");
        return alignment;
    }

    private static IWorkCellFill? ReadLegacyFill(XElement? property, IWorkLegacyXmlCatalog catalog) {
        if (property == null) return null;
        if (property.Element(Sf + "null") != null) return new IWorkCellFill(null);
        IWorkColor? color = catalog.Color(property.Elements().FirstOrDefault());
        if (color != null) return new IWorkCellFill(color);
        catalog.AddLoss("IWORK_LEGACY_TABLE_FILL_OMITTED", "An unsupported Numbers '09 cell fill was omitted.");
        return null;
    }

    private static IWorkNumberFormat? ReadLegacyComputedFormat(XElement? record, LegacyCellStyle style,
        IWorkLegacyXmlCatalog catalog) {
        XElement? computed = record?.Element(Sf + "cf");
        return computed == null ? style.NumberFormat : ReadLegacyNumericFormat(catalog.Resolve(
            computed.Element(Sf + "number-format") ?? computed.Element(Sf + "number-format-ref"), "number-format"), catalog);
    }

    private static IWorkNumberFormat? ReadLegacyNumberFormat(XElement style, IWorkLegacyXmlCatalog catalog) {
        string? family = (string?)catalog.Property(style, "SFTCellStylePropertyFormatType")
            ?.Element(Sf + "number")?.Attribute(Sfa + "number");
        if (family is null or "255") return null;
        if (family == "260") return new(IWorkNumberFormatKind.Text, null, false, IWorkNegativeNumberStyle.Minus);
        if (family == "261") {
            XElement? date = catalog.Resolve(catalog.Property(style, "SFTCellStylePropertyDateTimeFormat")
                ?.Elements().FirstOrDefault(), "date-format");
            IWorkDateTimeFormat? format = IWorkDateTimeFormat.CreateQualified((string?)date?.Attribute(Sf + "fmt") ?? "");
            if (format != null) return new(IWorkNumberFormatKind.DateTime, null, false, IWorkNegativeNumberStyle.Minus,
                dateTimeFormat: format);
        } else if (family is "256" or "257" or "258") {
            return ReadLegacyNumericFormat(catalog.Resolve(catalog.Property(style, "SFTCellStylePropertyNumberFormat")
                ?.Elements().FirstOrDefault(), "number-format"), catalog);
        }
        catalog.AddLoss("IWORK_LEGACY_TABLE_NUMBER_FORMAT_OMITTED",
            "A Numbers '09 cell retains its typed value without an unsupported display format.");
        return null;
    }

    private static IWorkNumberFormat? ReadLegacyNumericFormat(XElement? format, IWorkLegacyXmlCatalog catalog) {
        string? type = (string?)format?.Attribute(Sf + "format-type");
        string? decimals = (string?)format?.Attribute(Sf + "format-decimal-places");
        string? negative = (string?)format?.Attribute(Sf + "format-negative-style");
        string? grouping = (string?)format?.Attribute(Sf + "format-show-thousands-separator");
        string? accounting = (string?)format?.Attribute(Sf + "format-use-accounting-style");
        string? currency = (string?)format?.Attribute(Sf + "format-currency-code");
        bool supported = type is "0" or "1" or "2"
            && int.TryParse(decimals, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int places)
            && places is >= -1 and <= 30
            && int.TryParse(negative ?? "0", out int negativeStyle) && negativeStyle is >= 0 and <= 3
            && grouping is null or "false" or "true" && accounting is null or "false" or "true";
        if (!supported || type == "1" && (currency?.Length != 3 || currency.Any(character => character is < 'A' or > 'Z'))) {
            catalog.AddLoss("IWORK_LEGACY_TABLE_NUMBER_FORMAT_OMITTED",
                "A Numbers '09 cell retains its typed value without an unsupported display format.");
            return null;
        }
        int count = int.Parse(decimals!, CultureInfo.InvariantCulture);
        return new(type == "1" ? IWorkNumberFormatKind.Currency : type == "2" ? IWorkNumberFormatKind.Percentage
            : IWorkNumberFormatKind.Number, count < 0 ? null : count, grouping == "true",
            (IWorkNegativeNumberStyle)int.Parse(negative ?? "0", CultureInfo.InvariantCulture),
            currencyCode: type == "1" ? currency : null, useAccountingStyle: type == "1" && accounting == "true");
    }
}
