using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Builds a preservation-aware append-only rewrite for loaded XLSB style tables.</summary>
    internal sealed class XlsbStylesheetRewritePlan {
        private const int BrtBeginFonts = 611;
        private const int BrtEndFonts = 612;
        private const int BrtBeginCellXfs = 617;
        private const int BrtEndCellXfs = 618;
        private const int MaximumFonts = 0xFFD3;
        private const int MaximumCellFormats = 0xFF96;

        private XlsbStylesheetRewritePlan(string? partName, byte[]? replacement, int cellFormatCount) {
            PartName = partName;
            Replacement = replacement;
            CellFormatCount = cellFormatCount;
        }

        internal string? PartName { get; }

        internal byte[]? Replacement { get; }

        internal int CellFormatCount { get; }

        internal static XlsbStylesheetRewritePlan Create(
            ExcelDocument document,
            XlsbWorkbook sourceWorkbook,
            byte[]? sourcePartBytes) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sourceWorkbook == null) throw new ArgumentNullException(nameof(sourceWorkbook));

            Stylesheet? current = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            if (sourceWorkbook.Stylesheet == null) {
                if (current != null) {
                    throw new NotSupportedException("Native XLSB rewriting cannot add a workbook style table to a source that did not contain one. Save style changes as .xlsx.");
                }
                return new XlsbStylesheetRewritePlan(partName: null, replacement: null, cellFormatCount: 1);
            }
            if (string.IsNullOrWhiteSpace(sourceWorkbook.StylesheetPartName) || sourcePartBytes == null) {
                throw new InvalidDataException("The loaded XLSB style table has no source package part.");
            }
            if (current == null) {
                throw new NotSupportedException("Native XLSB rewriting cannot remove the workbook style table. Save style changes as .xlsx.");
            }

            Stylesheet expected = XlsbStylesheetProjector.Create(sourceWorkbook.Stylesheet);
            if (string.Equals(current.OuterXml, expected.OuterXml, StringComparison.Ordinal)) {
                return new XlsbStylesheetRewritePlan(
                    sourceWorkbook.StylesheetPartName,
                    replacement: null,
                    sourceWorkbook.Stylesheet.CellFormats.Count);
            }

            ValidateRootAndUnchangedCollections(current, expected);
            Font[] addedFonts = GetAppendedItems<Font>(current.Fonts, expected.Fonts, "fonts");
            CellFormat[] addedCellFormats = GetAppendedItems<CellFormat>(current.CellFormats, expected.CellFormats, "cell formats");
            if (addedFonts.Length == 0 && addedCellFormats.Length == 0) {
                throw new NotSupportedException("Native XLSB rewriting found an unsupported workbook style-table mutation. Save style changes as .xlsx.");
            }

            int fontCount = checked(sourceWorkbook.Stylesheet.Fonts.Count + addedFonts.Length);
            int cellFormatCount = checked(sourceWorkbook.Stylesheet.CellFormats.Count + addedCellFormats.Length);
            if (fontCount > MaximumFonts) {
                throw new NotSupportedException($"Native XLSB rewriting supports at most {MaximumFonts:N0} fonts.");
            }
            if (cellFormatCount > MaximumCellFormats) {
                throw new NotSupportedException($"Native XLSB rewriting supports at most {MaximumCellFormats:N0} cell formats.");
            }
            ValidateAddedCellFormats(
                addedCellFormats,
                fontCount,
                sourceWorkbook.Stylesheet.Fills.Count,
                sourceWorkbook.Stylesheet.Borders.Count,
                sourceWorkbook.Stylesheet.CellStyleFormats.Count);

            XlsbGeneratedRecord[] fontRecords = addedFonts
                .Select(XlsbStylesheetPartWriter.CreateFontRecord)
                .ToArray();
            XlsbGeneratedRecord[] cellFormatRecords = addedCellFormats
                .Select(format => XlsbStylesheetPartWriter.CreateCellFormatRecord(format))
                .ToArray();
            byte[] replacement = RewriteCollections(
                sourcePartBytes,
                sourceWorkbook.Stylesheet.Fonts.Count,
                sourceWorkbook.Stylesheet.CellFormats.Count,
                fontRecords,
                cellFormatRecords);
            return new XlsbStylesheetRewritePlan(sourceWorkbook.StylesheetPartName, replacement, cellFormatCount);
        }

        private static void ValidateRootAndUnchangedCollections(Stylesheet current, Stylesheet expected) {
            if (!AttributesMatch(current, expected, ignoreCount: false)) ThrowUnsupportedMutation();
            OpenXmlElement[] currentChildren = current.ChildElements.ToArray();
            OpenXmlElement[] expectedChildren = expected.ChildElements.ToArray();
            if (currentChildren.Length != expectedChildren.Length) ThrowUnsupportedMutation();

            for (int index = 0; index < currentChildren.Length; index++) {
                OpenXmlElement actual = currentChildren[index];
                OpenXmlElement baseline = expectedChildren[index];
                if (actual.GetType() != baseline.GetType()) ThrowUnsupportedMutation();
                if (actual is Fonts || actual is CellFormats) {
                    if (!AttributesMatch(actual, baseline, ignoreCount: true)) ThrowUnsupportedMutation();
                    continue;
                }
                if (!string.Equals(actual.OuterXml, baseline.OuterXml, StringComparison.Ordinal)) ThrowUnsupportedMutation();
            }
        }

        private static T[] GetAppendedItems<T>(
            OpenXmlCompositeElement? current,
            OpenXmlCompositeElement? expected,
            string collectionName) where T : OpenXmlElement {
            T[] actualItems = current?.Elements<T>().ToArray() ?? Array.Empty<T>();
            T[] expectedItems = expected?.Elements<T>().ToArray() ?? Array.Empty<T>();
            if ((current != null && current.ChildElements.Count != actualItems.Length)
                || (expected != null && expected.ChildElements.Count != expectedItems.Length)) {
                throw new NotSupportedException($"Native XLSB rewriting found unsupported child content in the {collectionName} collection.");
            }
            if (actualItems.Length < expectedItems.Length) ThrowUnsupportedMutation();
            for (int index = 0; index < expectedItems.Length; index++) {
                if (!string.Equals(actualItems[index].OuterXml, expectedItems[index].OuterXml, StringComparison.Ordinal)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot modify or reorder existing {collectionName}. Save style changes as .xlsx.");
                }
            }
            return actualItems.Skip(expectedItems.Length).ToArray();
        }

        private static bool AttributesMatch(OpenXmlElement current, OpenXmlElement expected, bool ignoreCount) {
            OpenXmlAttribute[] actual = current.GetAttributes()
                .Where(attribute => !ignoreCount || !string.Equals(attribute.LocalName, "count", StringComparison.Ordinal))
                .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
                .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal)
                .ToArray();
            OpenXmlAttribute[] baseline = expected.GetAttributes()
                .Where(attribute => !ignoreCount || !string.Equals(attribute.LocalName, "count", StringComparison.Ordinal))
                .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
                .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal)
                .ToArray();
            return actual.Length == baseline.Length
                && actual.Zip(baseline, (left, right) =>
                    left.LocalName == right.LocalName
                    && left.NamespaceUri == right.NamespaceUri
                    && left.Value == right.Value).All(match => match);
        }

        private static void ValidateAddedCellFormats(
            IEnumerable<CellFormat> formats,
            int fontCount,
            int fillCount,
            int borderCount,
            int styleFormatCount) {
            foreach (CellFormat format in formats) {
                uint fontId = format.FontId?.Value ?? 0U;
                uint fillId = format.FillId?.Value ?? 0U;
                uint borderId = format.BorderId?.Value ?? 0U;
                uint parentId = format.FormatId?.Value ?? 0U;
                if (fontId >= fontCount || fillId >= fillCount || borderId >= borderCount || parentId >= styleFormatCount) {
                    throw new NotSupportedException("Native XLSB rewriting found an appended cell format with an out-of-range style reference.");
                }
            }
        }

        private static byte[] RewriteCollections(
            byte[] sourcePartBytes,
            int sourceFontCount,
            int sourceCellFormatCount,
            IReadOnlyList<XlsbGeneratedRecord> addedFonts,
            IReadOnlyList<XlsbGeneratedRecord> addedCellFormats) {
            IReadOnlyList<XlsbRecord> records;
            using (var input = new MemoryStream(sourcePartBytes, writable: false)) {
                records = XlsbRecordReader.ReadAll(input);
            }
            int beginFonts = FindSingleRecord(records, BrtBeginFonts, "BrtBeginFonts");
            int endFonts = FindSingleRecord(records, BrtEndFonts, "BrtEndFonts");
            int beginCellFormats = FindSingleRecord(records, BrtBeginCellXfs, "BrtBeginCellXFs");
            int endCellFormats = FindSingleRecord(records, BrtEndCellXfs, "BrtEndCellXFs");
            ValidateDeclaredCount(records[beginFonts], sourceFontCount, "font");
            ValidateDeclaredCount(records[beginCellFormats], sourceCellFormatCount, "cell format");

            using var output = new MemoryStream(sourcePartBytes.Length + (addedFonts.Count + addedCellFormats.Count) * 64);
            for (int index = 0; index < records.Count; index++) {
                XlsbRecord record = records[index];
                if (index == beginFonts) {
                    WriteCountRecord(output, record.Type, checked(sourceFontCount + addedFonts.Count));
                    continue;
                }
                if (index == beginCellFormats) {
                    WriteCountRecord(output, record.Type, checked(sourceCellFormatCount + addedCellFormats.Count));
                    continue;
                }
                if (index == endFonts) WriteGeneratedRecords(output, addedFonts);
                if (index == endCellFormats) WriteGeneratedRecords(output, addedCellFormats);
                XlsbRecordWriter.Write(output, record.Type, record.Data);
            }
            return output.ToArray();
        }

        private static int FindSingleRecord(IReadOnlyList<XlsbRecord> records, int type, string name) {
            int found = -1;
            for (int index = 0; index < records.Count; index++) {
                if (records[index].Type != type) continue;
                if (found >= 0) throw new InvalidDataException($"The XLSB styles part contains more than one {name} record.");
                found = index;
            }
            if (found < 0) throw new InvalidDataException($"The XLSB styles part is missing its {name} record.");
            return found;
        }

        private static void ValidateDeclaredCount(XlsbRecord record, int expected, string description) {
            if (record.Data.Length != 4) {
                throw new InvalidDataException($"The XLSB {description} collection has an invalid count payload.");
            }
            uint declared = (uint)(record.Data[0]
                | (record.Data[1] << 8)
                | (record.Data[2] << 16)
                | (record.Data[3] << 24));
            if (declared != expected) {
                throw new InvalidDataException($"The XLSB {description} collection declares {declared} items but the loaded model contains {expected}.");
            }
        }

        private static void WriteCountRecord(Stream output, int type, int count) {
            byte[] payload = {
                (byte)count,
                (byte)(count >> 8),
                (byte)(count >> 16),
                (byte)(count >> 24)
            };
            XlsbRecordWriter.Write(output, type, payload);
        }

        private static void WriteGeneratedRecords(Stream output, IEnumerable<XlsbGeneratedRecord> records) {
            foreach (XlsbGeneratedRecord record in records) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
        }

        private static void ThrowUnsupportedMutation() =>
            throw new NotSupportedException("Native XLSB rewriting supports only append-only font and cell-format additions to the workbook style table. Save other style changes as .xlsx.");
    }
}
