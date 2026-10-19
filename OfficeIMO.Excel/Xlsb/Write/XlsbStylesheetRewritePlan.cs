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
            byte[]? sourcePartBytes, long maximumBytes = int.MaxValue) {
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
            Font[] fonts = GetAppendedItems<Font>(current.Fonts, expected.Fonts, "fonts");
            Fill[] fills = GetAppendedItems<Fill>(current.Fills, expected.Fills, "fills");
            Border[] borders = GetAppendedItems<Border>(current.Borders, expected.Borders, "borders");
            NumberingFormat[] formats = GetAppendedItems<NumberingFormat>(current.NumberingFormats, expected.NumberingFormats, "number formats");
            CellFormat[] xfs = GetAppendedItems<CellFormat>(current.CellFormats, expected.CellFormats, "cell formats");
            int fontCount = checked(sourceWorkbook.Stylesheet.Fonts.Count + fonts.Length);
            int fillCount = checked(sourceWorkbook.Stylesheet.Fills.Count + fills.Length);
            int borderCount = checked(sourceWorkbook.Stylesheet.Borders.Count + borders.Length);
            int cellFormatCount = checked(sourceWorkbook.Stylesheet.CellFormats.Count + xfs.Length);
            if (fontCount > MaximumFonts || cellFormatCount > MaximumCellFormats || fillCount > ushort.MaxValue || borderCount > ushort.MaxValue)
                throw new NotSupportedException("The XLSB formatting collection exceeds its native index limit.");
            ValidateAddedCellFormats(xfs, fontCount, fillCount, borderCount, sourceWorkbook.Stylesheet.CellStyleFormats.Count);
            var ids = new HashSet<uint>();
            foreach (var format in current.NumberingFormats?.Elements<NumberingFormat>() ?? Enumerable.Empty<NumberingFormat>()) {
                uint id = format.NumberFormatId?.Value ?? 0;
                if (id < 164 || !ids.Add(id)) throw new NotSupportedException("Custom XLSB number formats require unique ids from 164 through 65535.");
            }
            if (xfs.Any(format => (format.NumberFormatId?.Value ?? 0) >= 164 && !ids.Contains(format.NumberFormatId!.Value)))
                throw new NotSupportedException("An appended XLSB cell format refers to an undefined custom number format.");
            var collections = new[] {
                new Collection(615, 616, sourceWorkbook.Stylesheet.NumberFormats.Count, formats.Select(XlsbStylesheetPartWriter.CreateNumberFormatRecord).ToArray()),
                new Collection(BrtBeginFonts, BrtEndFonts, sourceWorkbook.Stylesheet.Fonts.Count, fonts.Select(XlsbStylesheetPartWriter.CreateFontRecord).ToArray()),
                new Collection(603, 604, sourceWorkbook.Stylesheet.Fills.Count, fills.Select(XlsbStylesheetPartWriter.CreateFillRecord).ToArray()),
                new Collection(613, 614, sourceWorkbook.Stylesheet.Borders.Count, borders.Select(XlsbStylesheetPartWriter.CreateBorderRecord).ToArray()),
                new Collection(BrtBeginCellXfs, BrtEndCellXfs, sourceWorkbook.Stylesheet.CellFormats.Count, xfs.Select(format => XlsbStylesheetPartWriter.CreateCellFormatRecord(format)).ToArray())
            };
            byte[] replacement = RewriteCollections(sourcePartBytes, collections, maximumBytes);
            return new XlsbStylesheetRewritePlan(sourceWorkbook.StylesheetPartName, replacement, cellFormatCount);
        }

        private static void ValidateRootAndUnchangedCollections(Stylesheet current, Stylesheet expected) {
            if (!AttributesMatch(current, expected, ignoreCount: false)) ThrowUnsupportedMutation();
            static bool Editable(OpenXmlElement element) => element is Fonts or Fills or Borders or NumberingFormats or CellFormats;
            if (current.ChildElements.Where(Editable).GroupBy(element => element.GetType()).Any(group => group.Count() != 1)) ThrowUnsupportedMutation();
            OpenXmlElement[] currentChildren = current.ChildElements.Where(element => !Editable(element)).ToArray();
            OpenXmlElement[] expectedChildren = expected.ChildElements.Where(element => !Editable(element)).ToArray();
            if (currentChildren.Length != expectedChildren.Length) ThrowUnsupportedMutation();
            for (int index = 0; index < currentChildren.Length; index++)
                if (currentChildren[index].OuterXml != expectedChildren[index].OuterXml) ThrowUnsupportedMutation();
            foreach (var actual in current.ChildElements.Where(Editable)) {
                var baseline = expected.ChildElements.FirstOrDefault(element => element.GetType() == actual.GetType());
                if (baseline != null && !AttributesMatch(actual, baseline, ignoreCount: true)) ThrowUnsupportedMutation();
                if (baseline == null && actual.GetAttributes().Any(attribute => attribute.LocalName != "count")) ThrowUnsupportedMutation();
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

        private sealed record Collection(int Begin, int End, int OriginalCount, IReadOnlyList<XlsbGeneratedRecord> Added);

        private static byte[] RewriteCollections(byte[] sourcePartBytes, IReadOnlyList<Collection> collections, long maximumBytes) {
            using var input = new MemoryStream(sourcePartBytes, writable: false);
            var records = XlsbRecordReader.ReadAll(input);
            var missing = new HashSet<int>();
            foreach (var collection in collections) {
                var begins = records.Select((record, index) => (record, index)).Where(item => item.record.Type == collection.Begin).ToArray();
                var ends = records.Select((record, index) => (record, index)).Where(item => item.record.Type == collection.End).ToArray();
                if (begins.Length == 0 && ends.Length == 0 && collection.OriginalCount == 0) { missing.Add(collection.Begin); continue; }
                if (begins.Length != 1 || ends.Length != 1 || begins[0].index >= ends[0].index)
                    throw new InvalidDataException("The XLSB formatting collection has invalid framing.");
                ValidateDeclaredCount(begins[0].record, collection.OriginalCount, "formatting");
            }
            using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
            // Fmts is optional and precedes fonts. Other core collections are mandatory in a projected source.
            foreach (var record in records) {
                if (record.Type == BrtBeginFonts && missing.Contains(615)) {
                    var formats = collections[0];
                    if (formats.Added.Count > 0) {
                        WriteCountRecord(output, formats.Begin, formats.Added.Count);
                        WriteGeneratedRecords(output, formats.Added);
                        XlsbRecordWriter.Write(output, formats.End, Array.Empty<byte>());
                    }
                }
                var begin = collections.FirstOrDefault(collection => collection.Begin == record.Type);
                var end = collections.FirstOrDefault(collection => collection.End == record.Type);
                if (begin != null) WriteCountRecord(output, record.Type, checked(begin.OriginalCount + begin.Added.Count));
                else {
                    if (end != null) WriteGeneratedRecords(output, end.Added);
                    XlsbRecordWriter.Write(output, record.Type, record.Data);
                }
            }
            if (collections.Any(collection => collection.Begin != 615 && missing.Contains(collection.Begin) && collection.Added.Count > 0))
                throw new InvalidDataException("The source XLSB stylesheet is missing a required formatting collection.");
            return output.ToArray();
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
            throw new NotSupportedException("Native XLSB rewriting supports append-only font, fill, border, number-format and cell-format additions to the workbook style table. Save other style changes as .xlsx.");
    }
}
