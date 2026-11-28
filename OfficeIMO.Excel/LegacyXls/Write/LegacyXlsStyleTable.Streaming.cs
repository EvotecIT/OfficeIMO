using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal sealed partial class LegacyXlsStyleTable {
    internal const int StreamingStyleCount = 5;

    // Binary row writers use only the generated plain, date and duration styles.
    internal static LegacyXlsStyleTable CreateStreaming(OfficeXmlNode stylesheet, ushort fontIndex) {
        const int main = OfficeXmlNamespaces.SpreadsheetML;
        var numberFormats = new Dictionary<uint, string>();
        foreach (OfficeXmlNode format in stylesheet.Child(main, "numFmts").Children(main, "numFmt")) {
            if (format.UInt32Attribute("numFmtId") is uint id && format.Attribute("formatCode") is string code) {
                numberFormats.Add(id, code);
            }
        }

        var map = new Dictionary<uint, ushort>(StreamingStyleCount);
        var formats = new Dictionary<ushort, byte[]>();
        var records = new List<byte[]>(BuildRequiredStyleXfPayloads());
        uint index = 0;
        foreach (OfficeXmlNode format in stylesheet.Child(main, "cellXfs").Children(main, "xf")) {
            ushort numberFormat = ResolveNumberFormatId(index, format.UInt32Attribute("numFmtId") ?? 0U, numberFormats, formats);
            map.Add(index, checked((ushort)records.Count));
            records.Add(BuildXfPayload(numberFormat, fontIndex, fontIndex != 0 || format.BooleanAttribute("applyFont") == true,
                FillInfo.None, AlignmentInfo.None, BorderInfo.None, ProtectionInfo.Default));
            if (++index == StreamingStyleCount) break;
        }

        if (index != StreamingStyleCount) throw new InvalidDataException("The generated XLS stylesheet is missing one or more streaming cell styles.");
        return new LegacyXlsStyleTable(map, formats.OrderBy(pair => pair.Key).Select(pair => pair.Value).ToArray(),
            records, [BuildNormalStylePayload()]);
    }
}
