using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Excel.Xlsb.Projection;
using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Rewrites supported worksheet cells and hyperlinks in an existing XLSB package while copying every other part.
    /// </summary>
    internal static class XlsbNativePackageWriter {
        private const int MaxWorksheetPartBytes = 128 * 1024 * 1024;
        private const string HyperlinkRelationshipSuffix = "/hyperlink";
        private static readonly DateTimeOffset ReproducibleEntryTime =
            new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

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
            var deletions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var packageStream = new MemoryStream(sourcePackageBytes, writable: false))
            using (var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: false)) {
                byte[]? sourceStylesPart = null;
                if (!string.IsNullOrWhiteSpace(sourceWorkbook.StylesheetPartName)) {
                    ZipArchiveEntry sourceStylesEntry = FindEntry(archive, sourceWorkbook.StylesheetPartName!)
                        ?? throw new InvalidDataException($"The source XLSB styles part '{sourceWorkbook.StylesheetPartName}' is missing.");
                    sourceStylesPart = ReadEntry(sourceStylesEntry, sourceWorkbook.MaxPartBytes);
                }
                XlsbStylesheetRewritePlan stylesheetPlan = XlsbStylesheetRewritePlan.Create(
                    document,
                    sourceWorkbook,
                    sourceStylesPart);
                if (stylesheetPlan.Replacement != null) {
                    replacements.Add(stylesheetPlan.PartName!, stylesheetPlan.Replacement);
                }

                for (int index = 0; index < sheets.Length; index++) {
                    XlsbWorksheet sourceSheet = sourceWorkbook.Worksheets[index];
                    string partName = sourceSheet.PartName
                        ?? throw new InvalidDataException($"The source XLSB worksheet '{sourceSheet.Name}' has no resolved package part.");
                    ZipArchiveEntry sourceEntry = FindEntry(archive, partName)
                        ?? throw new InvalidDataException($"The source XLSB worksheet part '{partName}' is missing.");
                    byte[] originalPart = ReadEntry(sourceEntry, MaxWorksheetPartBytes);
                    IReadOnlyList<XlsbWriteCell> cells = XlsbWorksheetCellExtractor.Extract(
                        document,
                        sheets[index],
                        sourceSheet,
                        stylesheetPlan.CellFormatCount);
                    bool rewriteHyperlinks = !XlsbWorksheetHyperlinkProjector.Matches(sheets[index], sourceSheet);
                    string[] reservedRelationshipIds = sourceSheet.Relationships.Values
                        .Where(relationship => !relationship.Type.EndsWith(HyperlinkRelationshipSuffix, StringComparison.Ordinal))
                        .Select(relationship => relationship.Id)
                        .ToArray();
                    XlsbWorksheetHyperlinkPlan? hyperlinkPlan = rewriteHyperlinks
                        ? XlsbWorksheetHyperlinkPlan.Create(
                            sheets[index],
                            pruneOrphanedRelationships: true,
                            reservedRelationshipIds)
                        : null;
                    byte[] rewrittenPart = XlsbWorksheetPartWriter.Rewrite(
                        originalPart,
                        cells,
                        hyperlinkPlan?.Records ?? Array.Empty<XlsbGeneratedRecord>(),
                        rewriteHyperlinks);
                    if (!originalPart.SequenceEqual(rewrittenPart)) {
                        replacements.Add(partName, rewrittenPart);
                    }
                    if (hyperlinkPlan != null && !hyperlinkPlan.RelationshipsMatch(sourceSheet)) {
                        string relationshipPartName = GetRelationshipPartName(partName);
                        int preservedRelationshipCount = sourceSheet.Relationships.Values.Count(relationship =>
                            !relationship.Type.EndsWith(HyperlinkRelationshipSuffix, StringComparison.Ordinal));
                        if (hyperlinkPlan.Relationships.Count == 0 && preservedRelationshipCount == 0) {
                            deletions.Add(relationshipPartName);
                        } else {
                            replacements[relationshipPartName] = hyperlinkPlan.CreateRelationshipPart(sourceSheet.Relationships);
                        }
                    }
                }
            }

            if (replacements.Count == 0 && deletions.Count == 0) return sourcePackageBytes;
            byte[] rewritten = RewritePackage(
                sourcePackageBytes,
                replacements,
                sourceWorkbook.MaxPartBytes,
                sourceWorkbook.MaxPackageBytes,
                deletions);
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
                .FirstOrDefault(element => element is not Sheets
                    && element is not WorkbookProperties
                    && element is not WorkbookProtection
                    && element is not DefinedNames
                    && element is not CalculationProperties);
            if (unsupportedWorkbookChild != null) {
                throw new NotSupportedException($"Native XLSB rewriting currently accepts cell-value edits only. Workbook metadata '{unsupportedWorkbookChild.LocalName}' was modified; save as .xlsx.");
            }

            ValidateWorkbookProperties(document, sourceWorkbook);
            ValidateWorkbookProtection(document, sourceWorkbook);
            ValidateDefinedNames(document, sourceWorkbook);
            ValidateCalculationProperties(document, sourceWorkbook);
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

        private static void ValidateCalculationProperties(ExcelDocument document, XlsbWorkbook sourceWorkbook) {
            CalculationProperties? properties = document.WorkbookRoot.GetFirstChild<CalculationProperties>();
            if (!XlsbCalculationPropertiesProjector.Matches(properties, sourceWorkbook.CalculationProperties)) {
                throw new NotSupportedException("Native XLSB rewriting preserves but cannot modify workbook calculation properties. Save that change as .xlsx.");
            }
        }

        private static void ValidateWorkbookProtection(ExcelDocument document, XlsbWorkbook sourceWorkbook) {
            WorkbookProtection? protection = document.WorkbookRoot.GetFirstChild<WorkbookProtection>();
            if (!XlsbWorkbookProtectionProjector.Matches(protection, sourceWorkbook.WorkbookProtection)) {
                throw new NotSupportedException("Native XLSB rewriting preserves but cannot modify workbook protection. Save that change as .xlsx.");
            }
        }

        private static void ValidateDefinedNames(ExcelDocument document, XlsbWorkbook sourceWorkbook) {
            DefinedNames? definedNames = document.WorkbookRoot.GetFirstChild<DefinedNames>();
            if (!XlsbDefinedNameProjector.Matches(definedNames, sourceWorkbook.DefinedNames)) {
                throw new NotSupportedException("Native XLSB rewriting preserves but cannot modify workbook defined names. Save that change as .xlsx.");
            }
        }

        internal static byte[] RewritePackage(
            byte[] sourcePackageBytes,
            IReadOnlyDictionary<string, byte[]> replacements,
            int maxPartBytes,
            long maxPackageBytes) => RewritePackage(
                sourcePackageBytes,
                replacements,
                maxPartBytes,
                maxPackageBytes,
                deletions: null);

        internal static byte[] RewritePackage(
            byte[] sourcePackageBytes,
            IReadOnlyDictionary<string, byte[]> replacements,
            int maxPartBytes,
            long maxPackageBytes,
            ISet<string>? deletions) {
            if (sourcePackageBytes == null) throw new ArgumentNullException(nameof(sourcePackageBytes));
            if (replacements == null) throw new ArgumentNullException(nameof(replacements));

            var normalizedReplacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, byte[]> replacement in replacements) {
                string normalizedName = NormalizePartName(replacement.Key);
                if (replacement.Value == null) {
                    throw new ArgumentException($"Replacement package part '{normalizedName}' has no content.", nameof(replacements));
                }
                if (normalizedReplacements.ContainsKey(normalizedName)) {
                    throw new ArgumentException($"Replacement package part '{normalizedName}' is duplicated.", nameof(replacements));
                }
                normalizedReplacements.Add(normalizedName, replacement.Value);
            }
            var normalizedDeletions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (deletions != null) {
                foreach (string deletion in deletions) {
                    string normalizedName = NormalizePartName(deletion);
                    if (normalizedReplacements.ContainsKey(normalizedName)) {
                        throw new ArgumentException($"Package part '{normalizedName}' cannot be replaced and deleted in the same rewrite.", nameof(deletions));
                    }
                    normalizedDeletions.Add(normalizedName);
                }
            }

            using var sourceStream = new MemoryStream(sourcePackageBytes, writable: false);
            using var sourceArchive = new ZipArchive(sourceStream, ZipArchiveMode.Read, leaveOpen: false);
            using var destinationStream = new MemoryStream(sourcePackageBytes.Length + 4096);
            long decompressedBytes = 0;
            byte[] buffer = new byte[81920];
            var writtenReplacements = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var destinationArchive = new ZipArchive(destinationStream, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (ZipArchiveEntry sourceEntry in sourceArchive.Entries) {
                    if (string.IsNullOrEmpty(sourceEntry.Name)) continue;
                    string normalizedName = NormalizePartName(sourceEntry.FullName);
                    if (normalizedDeletions.Contains(normalizedName)) continue;
                    ZipArchiveEntry destinationEntry = destinationArchive.CreateEntry(normalizedName, CompressionLevel.Optimal);
                    try { destinationEntry.LastWriteTime = sourceEntry.LastWriteTime; } catch (ArgumentOutOfRangeException) { }
                    using Stream output = destinationEntry.Open();
                    if (normalizedReplacements.TryGetValue(normalizedName, out byte[]? replacement)) {
                        writtenReplacements.Add(normalizedName);
                        ChargeRewriteBytes(
                            normalizedName,
                            replacement.Length,
                            maxPartBytes,
                            maxPackageBytes,
                            ref decompressedBytes);
                        output.Write(replacement, 0, replacement.Length);
                    } else {
                        if (sourceEntry.Length > maxPartBytes) {
                            throw new InvalidDataException(
                                $"The XLSB package part '{normalizedName}' declares {sourceEntry.Length} decompressed bytes, exceeding the configured rewrite limit of {maxPartBytes} bytes.");
                        }
                        using Stream input = sourceEntry.Open();
                        int partBytes = 0;
                        while (true) {
                            int read = input.Read(buffer, 0, buffer.Length);
                            if (read == 0) break;
                            ChargeRewriteBytes(
                                normalizedName,
                                read,
                                maxPartBytes,
                                maxPackageBytes,
                                ref decompressedBytes,
                                ref partBytes);
                            output.Write(buffer, 0, read);
                        }
                    }
                }

                foreach (KeyValuePair<string, byte[]> replacement in normalizedReplacements) {
                    if (writtenReplacements.Contains(replacement.Key)) continue;
                    ZipArchiveEntry destinationEntry = destinationArchive.CreateEntry(replacement.Key, CompressionLevel.Optimal);
                    destinationEntry.LastWriteTime = ReproducibleEntryTime;
                    ChargeRewriteBytes(
                        replacement.Key,
                        replacement.Value.Length,
                        maxPartBytes,
                        maxPackageBytes,
                        ref decompressedBytes);
                    using Stream output = destinationEntry.Open();
                    output.Write(replacement.Value, 0, replacement.Value.Length);
                }
            }

            return destinationStream.ToArray();
        }

        private static string GetRelationshipPartName(string sourcePartName) {
            string source = NormalizePartName(sourcePartName);
            int separator = source.LastIndexOf('/');
            string directory = separator < 0 ? string.Empty : source.Substring(0, separator + 1);
            string fileName = separator < 0 ? source : source.Substring(separator + 1);
            return directory + "_rels/" + fileName + ".rels";
        }

        private static string NormalizePartName(string partName) {
            if (string.IsNullOrWhiteSpace(partName)) throw new ArgumentException("Package part name cannot be empty.", nameof(partName));
            string normalized = partName.Replace('\\', '/').TrimStart('/');
            if (normalized.EndsWith("/", StringComparison.Ordinal)
                || normalized.Split('/').Any(segment => segment.Length == 0 || segment == "." || segment == "..")) {
                throw new InvalidDataException($"The package part name '{partName}' is not safe.");
            }
            return normalized;
        }

        private static void ChargeRewriteBytes(
            string partName,
            int bytes,
            int maxPartBytes,
            long maxPackageBytes,
            ref long decompressedBytes) {
            int partBytes = 0;
            ChargeRewriteBytes(
                partName,
                bytes,
                maxPartBytes,
                maxPackageBytes,
                ref decompressedBytes,
                ref partBytes);
        }

        private static void ChargeRewriteBytes(
            string partName,
            int bytes,
            int maxPartBytes,
            long maxPackageBytes,
            ref long decompressedBytes,
            ref int partBytes) {
            if (bytes < 0 || partBytes > maxPartBytes - bytes) {
                throw new InvalidDataException(
                    $"The XLSB package part '{partName}' exceeds the configured rewrite limit of {maxPartBytes} bytes while decompressing.");
            }
            if (decompressedBytes > maxPackageBytes - bytes) {
                throw new InvalidDataException(
                    $"The XLSB package exceeds the configured aggregate rewrite limit of {maxPackageBytes} bytes while decompressing.");
            }

            partBytes += bytes;
            decompressedBytes += bytes;
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
