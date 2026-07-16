using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Projection;
using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Rewrites the supported cell-value subset of an existing XLSB package while copying every other part.
    /// </summary>
    internal static class XlsbNativePackageWriter {
        private const int MaxWorksheetPartBytes = 128 * 1024 * 1024;

        internal static byte[] Rewrite(
            ExcelDocument document,
            byte[] sourcePackageBytes,
            XlsbWorkbook sourceWorkbook) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sourcePackageBytes == null) throw new ArgumentNullException(nameof(sourcePackageBytes));
            if (sourceWorkbook == null) throw new ArgumentNullException(nameof(sourceWorkbook));

            ThrowIfUnsupportedWorkbookMutation(document, sourceWorkbook, sourcePackageBytes);
            ExcelSheet[] sheets = document.Sheets.ToArray();
            var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            using (var packageStream = new MemoryStream(sourcePackageBytes, writable: false))
            using (var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false)) {
                for (int index = 0; index < sheets.Length; index++) {
                    XlsbWorksheet sourceSheet = sourceWorkbook.Worksheets[index];
                    string partName = sourceSheet.PartName
                        ?? throw new InvalidDataException($"The source XLSB worksheet '{sourceSheet.Name}' has no resolved package part.");
                    ZipArchiveEntry sourceEntry = FindEntry(archive, partName)
                        ?? throw new InvalidDataException($"The source XLSB worksheet part '{partName}' is missing.");
                    byte[] originalPart = ReadEntry(sourceEntry, MaxWorksheetPartBytes);
                    IReadOnlyList<XlsbWriteCell> cells = XlsbWorksheetCellExtractor.Extract(document, sheets[index], sourceSheet);
                    byte[] rewrittenPart = XlsbWorksheetPartWriter.Rewrite(originalPart, cells);
                    if (!originalPart.SequenceEqual(rewrittenPart)) {
                        replacements.Add(partName, rewrittenPart);
                    }
                }
            }

            if (replacements.Count == 0) return sourcePackageBytes;
            byte[] rewritten = RewritePackage(sourcePackageBytes, replacements);
            if (!XlsbPackageDetector.TryFindWorkbookPart(rewritten, out _)) {
                throw new InvalidDataException("The rewritten package no longer satisfies the XLSB package contract.");
            }

            // Re-read the result before exposing it so record framing, relationships, and projected cells are proven.
            XlsbWorkbookReader.Load(rewritten, new XlsbImportOptions { ReportPreservedRecords = false });
            return rewritten;
        }

        private static void ThrowIfUnsupportedWorkbookMutation(
            ExcelDocument document,
            XlsbWorkbook sourceWorkbook,
            byte[] sourcePackageBytes) {
            if (document.HasPackagePropertiesDirty) {
                throw new NotSupportedException("Native XLSB rewriting currently accepts cell-value edits only. Document-property changes must be saved as .xlsx.");
            }

            OpenXmlElement? unsupportedWorkbookChild = document.WorkbookRoot.ChildElements
                .FirstOrDefault(element => element is not Sheets && element is not WorkbookProperties);
            if (unsupportedWorkbookChild != null) {
                throw new NotSupportedException($"Native XLSB rewriting currently accepts cell-value edits only. Workbook metadata '{unsupportedWorkbookChild.LocalName}' was modified; save as .xlsx.");
            }

            ValidateWorkbookProperties(document, sourceWorkbook);
            ValidateStylesheet(document, sourceWorkbook);

            ExcelSheet[] sheets = document.Sheets.ToArray();
            if (sheets.Length != sourceWorkbook.Worksheets.Count) {
                throw new NotSupportedException("Native XLSB rewriting currently requires the original worksheet set and order. Save workbook structure changes as .xlsx.");
            }

            for (int index = 0; index < sheets.Length; index++) {
                XlsbWorksheet sourceSheet = sourceWorkbook.Worksheets[index];
                ExcelSheet currentSheet = sheets[index];
                uint currentState = currentSheet.VeryHidden ? 2U : currentSheet.Hidden ? 1U : 0U;
                if (!string.Equals(currentSheet.Name, sourceSheet.Name, StringComparison.Ordinal)
                    || currentState != sourceSheet.State) {
                    throw new NotSupportedException("Native XLSB rewriting currently requires original worksheet names, order, and visibility. Save workbook structure changes as .xlsx.");
                }
            }

            using var packageStream = new MemoryStream(sourcePackageBytes, writable: false);
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false);
            if (archive.Entries.Any(entry => entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase))) {
                throw new NotSupportedException("Native XLSB rewriting is blocked because the source package is digitally signed. Rewriting worksheet parts would invalidate the signature.");
            }
        }

        private static void ValidateWorkbookProperties(ExcelDocument document, XlsbWorkbook sourceWorkbook) {
            WorkbookProperties? properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            if (properties == null) {
                if (sourceWorkbook.Uses1904DateSystem) {
                    throw new NotSupportedException("Native XLSB rewriting cannot change the workbook date system. Save the workbook as .xlsx.");
                }
                return;
            }

            bool hasOnlyProjectedDateSystem = !properties.HasChildren
                && properties.GetAttributes().All(attribute =>
                    string.Equals(attribute.LocalName, "date1904", StringComparison.Ordinal)
                    && string.Equals(attribute.NamespaceUri, string.Empty, StringComparison.Ordinal));
            bool uses1904 = properties.Date1904?.Value == true;
            if (!hasOnlyProjectedDateSystem || uses1904 != sourceWorkbook.Uses1904DateSystem) {
                throw new NotSupportedException("Native XLSB rewriting currently cannot change workbook properties or the workbook date system. Save the workbook as .xlsx.");
            }
        }

        private static void ValidateStylesheet(ExcelDocument document, XlsbWorkbook sourceWorkbook) {
            if (sourceWorkbook.Stylesheet == null) return;

            Stylesheet? current = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
            Stylesheet expected = XlsbStylesheetProjector.Create(sourceWorkbook.Stylesheet);
            if (current == null || !string.Equals(current.OuterXml, expected.OuterXml, StringComparison.Ordinal)) {
                throw new NotSupportedException("Native XLSB rewriting currently preserves but cannot modify the workbook style table. Save style changes as .xlsx.");
            }
        }

        private static byte[] RewritePackage(byte[] sourcePackageBytes, IReadOnlyDictionary<string, byte[]> replacements) {
            using var sourceStream = new MemoryStream(sourcePackageBytes, writable: false);
            using var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false);
            using var destinationStream = new MemoryStream(sourcePackageBytes.Length + 4096);
            using (var destinationArchive = new ZipArchive(destinationStream, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (ZipArchiveEntry sourceEntry in sourceArchive.Entries) {
                    string normalizedName = sourceEntry.FullName.Replace('\\', '/');
                    ZipArchiveEntry destinationEntry = destinationArchive.CreateEntry(normalizedName, CompressionLevel.Optimal);
                    try { destinationEntry.LastWriteTime = sourceEntry.LastWriteTime; } catch (ArgumentOutOfRangeException) { }
                    using Stream output = destinationEntry.Open();
                    if (replacements.TryGetValue(normalizedName, out byte[]? replacement)) {
                        output.Write(replacement, 0, replacement.Length);
                    } else {
                        using Stream input = sourceEntry.Open();
                        input.CopyTo(output);
                    }
                }
            }

            return destinationStream.ToArray();
        }

        private static byte[] ReadEntry(ZipArchiveEntry entry, int maxBytes) {
            if (entry.Length > maxBytes) {
                throw new InvalidDataException($"The XLSB worksheet part '{entry.FullName}' exceeds the supported rewrite limit of {maxBytes} bytes.");
            }

            using Stream input = entry.Open();
            using var output = new MemoryStream(checked((int)entry.Length));
            byte[] buffer = new byte[81920];
            while (true) {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (output.Length + read > maxBytes) {
                    throw new InvalidDataException($"The XLSB worksheet part '{entry.FullName}' exceeds the supported rewrite limit of {maxBytes} bytes while decompressing.");
                }
                output.Write(buffer, 0, read);
            }

            return output.ToArray();
        }

        private static ZipArchiveEntry? FindEntry(ZipArchive archive, string partName) {
            return archive.Entries.FirstOrDefault(entry =>
                string.Equals(entry.FullName.Replace('\\', '/'), partName.Replace('\\', '/'), StringComparison.OrdinalIgnoreCase));
        }
    }
}
