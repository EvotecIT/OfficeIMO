using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    private const int Sml = OfficeXmlNamespaces.SpreadsheetML;

    private string?[]? ReadBoundedSourceLinkedStyles() {
        OfficeXmlNode stylesheet = ExcelPartRef.Workbook(_document).WorkbookStyles.Root;
        if (stylesheet.IsNone) return new[] { "General" };
        int remaining = 100_000;
        var custom = new Dictionary<uint, string>();
        var styles = new List<string?>();
        var namedStyles = new List<uint>();
        foreach (OfficeXmlNode item in stylesheet.Child(Sml, "numFmts").Children(Sml, "numFmt")) {
            if (--remaining < 0) return null;
            if (!uint.TryParse(item.Attribute("numFmtId"), NumberStyles.None, CultureInfo.InvariantCulture, out uint id) || item.Attribute("formatCode") is not string code || custom.ContainsKey(id)) return null;
            custom.Add(id, code);
        }
        foreach (OfficeXmlNode style in stylesheet.Child(Sml, "cellStyleXfs").Children(Sml, "xf")) {
            if (--remaining < 0 || !uint.TryParse(style.Attribute("numFmtId") ?? "0", NumberStyles.None, CultureInfo.InvariantCulture, out uint id)) return null;
            namedStyles.Add(id);
        }
        foreach (OfficeXmlNode style in stylesheet.Child(Sml, "cellXfs").Children(Sml, "xf")) {
            if (--remaining < 0) return null;
            string? formatId = style.Attribute("xfId");
            string? numberFormatId = style.Attribute("numFmtId");
            string? applyNumberFormat = style.Attribute("applyNumberFormat");
            uint baseIndex = 0;
            if (formatId != null && !uint.TryParse(formatId, NumberStyles.None, CultureInfo.InvariantCulture, out baseIndex)) return null;
            if (formatId != null && baseIndex >= namedStyles.Count) return null;
            uint id = formatId != null ? namedStyles[(int)baseIndex] : 0;
            // An applyNumberFormat value that is not a boolean makes the stylesheet unreadable here.
            bool? apply = OfficeXmlViewValues.ParseBoolean(applyNumberFormat);
            if (applyNumberFormat != null && apply == null) return null;
            if (apply != false && numberFormatId != null &&
                !uint.TryParse(numberFormatId, NumberStyles.None, CultureInfo.InvariantCulture, out id)) return null;
            styles.Add(custom.TryGetValue(id, out string? code) ? code : ExcelBuiltInNumberFormats.GetCode(id));
        }
        return styles.Count == 0 ? new[] { "General" } : styles.ToArray();
    }
}
