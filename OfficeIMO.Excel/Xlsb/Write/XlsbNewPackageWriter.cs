using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Creates a new, first-party XLSB package from the supported workbook subset.</summary>
    internal static class XlsbNewPackageWriter {
        private static readonly DateTimeOffset ReproducibleEntryTime =
            new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

        internal static void Write(ExcelDocument document, Stream destination) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("The XLSB destination must be writable.", nameof(destination));

            ExcelSheet[] sheets = document.Sheets.ToArray();
            ValidateWorkbook(document, sheets);
            Stylesheet? stylesheet = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            byte[]? stylesPart = null;
            int cellFormatCount = 1;
            if (stylesheet != null) {
                stylesPart = XlsbStylesheetPartWriter.Create(stylesheet, out cellFormatCount);
            }
            var worksheetParts = new byte[sheets.Length][];
            for (int index = 0; index < sheets.Length; index++) {
                IReadOnlyList<XlsbWriteCell> cells = XlsbWorksheetCellExtractor.ExtractNew(document, sheets[index]);
                XlsbWriteCell? invalidStyle = cells.FirstOrDefault(cell => cell.StyleIndex >= cellFormatCount);
                if (invalidStyle != null) {
                    throw new NotSupportedException($"Native XLSB generation found cell {sheets[index].Name}!R{invalidStyle.Row}C{invalidStyle.Column} with missing style index {invalidStyle.StyleIndex}.");
                }
                worksheetParts[index] = XlsbWorksheetPartWriter.Create(cells);
            }

            using var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
            WriteEntry(archive, "[Content_Types].xml", CreateContentTypes(sheets.Length, stylesPart != null));
            WriteEntry(archive, "_rels/.rels", RootRelationships);
            WriteEntry(archive, "xl/workbook.bin", XlsbWorkbookPartWriter.Create(
                sheets,
                document.DateSystem == ExcelDateSystem.NineteenFour));
            WriteEntry(archive, "xl/_rels/workbook.bin.rels", CreateWorkbookRelationships(sheets.Length, stylesPart != null));
            for (int index = 0; index < worksheetParts.Length; index++) {
                WriteEntry(
                    archive,
                    "xl/worksheets/sheet" + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin",
                    worksheetParts[index]);
            }
            if (stylesPart != null) WriteEntry(archive, "xl/styles.bin", stylesPart);
        }

        private static void ValidateWorkbook(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            if (sheets.Count == 0) {
                throw new NotSupportedException("Native XLSB generation requires at least one worksheet.");
            }

            if (document.HasPackagePropertiesDirty) {
                throw new NotSupportedException("Native XLSB generation does not yet support modified document properties.");
            }

            OpenXmlElement? unsupportedWorkbookChild = document.WorkbookRoot.ChildElements
                .FirstOrDefault(element => element is not Sheets && element is not WorkbookProperties);
            if (unsupportedWorkbookChild != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook metadata '{unsupportedWorkbookChild.LocalName}'.");
            }

            WorkbookProperties? properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            if (properties != null) {
                bool hasOnlyDateSystem = !properties.HasChildren
                    && properties.GetAttributes().All(attribute =>
                        string.Equals(attribute.LocalName, "date1904", StringComparison.Ordinal)
                        && string.Equals(attribute.NamespaceUri, string.Empty, StringComparison.Ordinal));
                if (!hasOnlyDateSystem) {
                    throw new NotSupportedException("Native XLSB generation currently supports only the workbook date1904 property.");
                }
            }

            if (document.WorkbookPartRoot.ExternalRelationships.Any()) {
                throw new NotSupportedException("Native XLSB generation does not yet support external workbook relationships.");
            }

            OpenXmlPart? unsupportedPart = document.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .FirstOrDefault(part => part is not WorksheetPart
                    && part is not SharedStringTablePart
                    && part is not WorkbookStylesPart);
            if (unsupportedPart != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook part '{unsupportedPart.ContentType}'.");
            }
        }

        private static string CreateContentTypes(int worksheetCount, bool hasStyles) {
            var builder = new StringBuilder(512 + worksheetCount * 120);
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            builder.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
            builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            for (int index = 0; index < worksheetCount; index++) {
                builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                builder.Append((index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(".bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>");
            }
            if (hasStyles) {
                builder.Append("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>");
            }
            builder.Append("</Types>");
            return builder.ToString();
        }

        private static string CreateWorkbookRelationships(int worksheetCount, bool hasStyles) {
            var builder = new StringBuilder(256 + worksheetCount * 180);
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int index = 0; index < worksheetCount; index++) {
                string number = (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                builder.Append("<Relationship Id=\"rId");
                builder.Append(number);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet");
                builder.Append(number);
                builder.Append(".bin\"/>");
            }
            if (hasStyles) {
                builder.Append("<Relationship Id=\"rId");
                builder.Append((worksheetCount + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.bin\"/>");
            }
            builder.Append("</Relationships>");
            return builder.ToString();
        }

        private static void WriteEntry(ZipArchive archive, string name, string content) {
            WriteEntry(archive, name, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content));
        }

        private static void WriteEntry(ZipArchive archive, string name, byte[] content) {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
            entry.LastWriteTime = ReproducibleEntryTime;
            using Stream output = entry.Open();
            output.Write(content, 0, content.Length);
        }

        private const string RootRelationships =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "</Relationships>";
    }
}
