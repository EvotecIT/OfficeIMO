using OfficeIMO.Excel.Xlsb.Projection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO.Compression;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Creates a new, first-party XLSB package from the supported workbook subset.</summary>
    internal static partial class XlsbNewPackageWriter {
        private static readonly DateTimeOffset ReproducibleEntryTime =
            new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

        internal static void Write(
            ExcelDocument document,
            Stream destination,
            bool useSharedStrings = false,
            CancellationToken cancellationToken = default, long maximumBytes = int.MaxValue) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(destination);
            if (!destination.CanWrite) throw new ArgumentException("The XLSB destination must be writable.", nameof(destination));

            ExcelSheet[] sheets = document.Sheets.ToArray();
            ValidateWorkbook(document, sheets);
            Stylesheet stylesheet = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet ?? ExcelDocument.CreateDefaultStylesheet();
            byte[] stylesPart = XlsbStylesheetPartWriter.CreateForNewWorkbook(stylesheet, out int cellFormatCount, maximumBytes);
            var richText = new XlsbRichTextFontTable(stylesheet,
                new XlsbCellTextRuns(document.WorkbookPartRoot.SharedStringTablePart?.SharedStringTable));
            long remainingBytes = maximumBytes - stylesPart.Length;
            var worksheetParts = new byte[sheets.Length][];
            var hyperlinkPlans = new XlsbWorksheetHyperlinkPlan[sheets.Length];
            var packageParts = new XlsbWorksheetPackageParts[sheets.Length];
            // Related parts copied from the workbook (drawings, charts, media) must not take the names of fixed parts.
            var partNames = new XlsbPartNameAllocator(Enumerable.Range(0, sheets.Length).Select(WorksheetPartName)
                .Concat(new[] { "xl/workbook.bin", "xl/styles.bin", "xl/sharedStrings.bin", "xl/theme/theme1.xml" }));
            XlsbSharedStringTable? sharedStrings = useSharedStrings ? new XlsbSharedStringTable() : null;
            var formulas = new XlsbNewWorkbookFormulaContext(document, sheets, XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbDefinedNameProjector.CollectionName));
            for (int index = 0; index < sheets.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<XlsbWriteCell> cells = XlsbWorksheetCellExtractor.ExtractNew(document, sheets[index], formulas, richText);
                hyperlinkPlans[index] = XlsbWorksheetHyperlinkPlan.Create(sheets[index]);
                packageParts[index] = new XlsbWorksheetPackageParts(WorksheetPartName(index), partNames,
                    hyperlinkPlans[index].Relationships.Select(relationship => relationship.Id));
                XlsbWorksheetFeatureRecords features = XlsbWorksheetFeatureRecords.Create(
                    new XlsbWorksheetFeatureContext(document, sheets[index], formulas, packageParts[index], richText));
                XlsbWriteCell? invalidStyle = cells.FirstOrDefault(cell => cell.StyleIndex >= cellFormatCount);
                if (invalidStyle != null) {
                    throw new NotSupportedException($"Native XLSB generation found cell {sheets[index].Name}!R{invalidStyle.Row}C{invalidStyle.Column} with missing style index {invalidStyle.StyleIndex}.");
                }
                worksheetParts[index] = XlsbWorksheetPartWriter.Create(
                    sheets[index],
                    cells,
                    cellFormatCount,
                    hyperlinkPlans[index].Records,
                    sharedStrings, Math.Max(1, remainingBytes), features: features);
                remainingBytes -= worksheetParts[index].LongLength;
                if (remainingBytes < 0) throw OfficeIMO.Core.Internal.OfficeOutputLimit.Create("Native XLSB intermediate parts exceed the in-memory save limit.");
            }

            if (richText.AddedFonts.Count != 0) {
                // Rich runs reference fonts appended after the workbook's own collection.
                remainingBytes += stylesPart.Length;
                stylesPart = XlsbStylesheetPartWriter.CreateForNewWorkbook(stylesheet, out _, Math.Max(1, remainingBytes), richText.AddedFonts);
                remainingBytes -= stylesPart.Length;
            }
            ArraySegment<byte>? sharedStringsPart = sharedStrings?.CreatePart(cancellationToken, Math.Max(1, remainingBytes));
            WritePackage(document, destination, sheets, stylesPart, worksheetParts, hyperlinkPlans, packageParts, sharedStringsPart, formulas);
        }

        private static string WorksheetPartName(int index) =>
            "xl/worksheets/sheet" + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".bin";

        internal static bool TryWriteDirectTabular(
            ExcelDocument document,
            ExcelDirectTabularSource source,
            Stream destination,
            CancellationToken cancellationToken,
            bool useSharedStrings = false,
            long maximumSharedStringBytes = ExcelSaveOptions.DefaultMaxInMemoryPackageBytes) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(destination);
            if (!destination.CanWrite) throw new ArgumentException("The XLSB destination must be writable.", nameof(destination));
            // Property edits are written by the general package writer.
            if (document.HasPackagePropertiesDirty || document.HasCustomDocumentPropertiesDirty) return false;

            ExcelSheet? sheet = document.GetOrCreateDeferredDirectTabularSheet(source.SheetName);
            if (sheet == null) {
                return false;
            }
            ExcelSheet[] sheets = [sheet];

            ValidateDirectTabularWorkbook(document, sheets, source.SheetName);
            cancellationToken.ThrowIfCancellationRequested();
            // Keep date cells and workbook metadata on the same policy for this save.
            ExcelDateSystem dateSystem = document.DateSystem;
            XlsbSharedStringTable? sharedStrings = useSharedStrings ? new XlsbSharedStringTable() : null;
            if (!XlsbDirectTabularPlan.CanCapture(source)) {
                return TryWriteStagedDirectTabularPackage(document, dateSystem, sheet, source, destination,
                    sharedStrings, maximumSharedStringBytes, cancellationToken);
            }

            // Freeze a flat source as a compressed package before publishing to an
            // ordinary memory destination. Other streams keep their existing path.
            if (destination.GetType() == typeof(MemoryStream)
                && destination.Length == 0 && destination.Position == 0
                && !cancellationToken.CanBeCanceled && document.Execution.OnTiming == null
                && ExcelDocument.HasFlatTabularSnapshot(source.Rows)) {
                return TryWriteStagedDirectTabularPackage(document, dateSystem, sheet, source, destination,
                    sharedStrings, maximumSharedStringBytes, cancellationToken, pooledSnapshot: true);
            }

            XlsbDirectTabularPlan? plan = null;
            try {
                if (!XlsbDirectTabularPlan.TryCreate(dateSystem, sheet, source, sharedStrings, cancellationToken, out plan)) return false;

                // Validate temporal styles and shared strings before touching the destination.
                Stylesheet? stylesheet = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
                byte[] stylesPart = XlsbStylesheetPartWriter.CreateForNewWorkbook(stylesheet, out _);
                ArraySegment<byte>? sharedStringsPart = sharedStrings?.CreatePart(cancellationToken, maximumSharedStringBytes);
                cancellationToken.ThrowIfCancellationRequested();
                if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
                WriteDirectTabularPackage(dateSystem, destination, source.SheetName, stylesPart,
                    sharedStringsPart, CreateThemeWriter(document.WorkbookPartRoot.ThemePart), plan, cancellationToken);
                if (destination.CanSeek) destination.SetLength(destination.Position);
                return true;
            } finally {
                plan?.Dispose();
            }
        }

        private static bool TryWriteStagedDirectTabularPackage(
            ExcelDocument document,
            ExcelDateSystem dateSystem,
            ExcelSheet sheet,
            ExcelDirectTabularSource source,
            Stream destination,
            XlsbSharedStringTable? sharedStrings,
            long maximumSharedStringBytes,
            CancellationToken cancellationToken,
            bool pooledSnapshot = false) {
            // Staging the compressed package preserves the read-once input contract
            // without retaining a complete uncompressed BIFF12 worksheet.
            using Stream package = pooledSnapshot
                ? new XlsbPackageBuffer()
                : new MemoryStream(XlsbWorksheetPartWriter.EstimateDirectWorksheetCapacity(source));
            using (ExcelPackageArchive archive = ExcelPackageArchive.Create(package, ReproducibleEntryTime.DateTime)) {
                using (Stream output = archive.OpenEntry("xl/worksheets/sheet1.bin", CompressionLevel.Fastest,
                    pooledSnapshot && source.Rows.RowCount >= 4096 && document.Execution.AllowsBackgroundProcessing)) {
                    if (!XlsbWorksheetPartWriter.TryWriteDirectTabular(
                        dateSystem, sheet, source, output, cancellationToken, sharedStrings)) return false;
                }

                // Dates can add styles and text can populate the shared-string table
                // while emitting rows. Finalize every dependent part before publication.
                Stylesheet? stylesheet = document.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet;
                byte[] stylesPart = XlsbStylesheetPartWriter.CreateForNewWorkbook(stylesheet, out _);
                ArraySegment<byte>? sharedStringsPart = sharedStrings?.CreatePart(cancellationToken, maximumSharedStringBytes);
                cancellationToken.ThrowIfCancellationRequested();
                WriteDirectTabularWorkbookEntries(archive, dateSystem, source.SheetName,
                    stylesPart != null, sharedStringsPart.HasValue, CreateThemeWriter(document.WorkbookPartRoot.ThemePart));
                if (stylesPart != null) archive.WriteEntry("xl/styles.bin", stylesPart, CompressionLevel.Fastest);
                if (sharedStringsPart.HasValue) archive.WriteEntry("xl/sharedStrings.bin", sharedStringsPart.Value, CompressionLevel.Fastest);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (pooledSnapshot && destination is MemoryStream memory && memory.Capacity < package.Length) {
                memory.Capacity = checked((int)package.Length);
            }
            if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
            if (package is XlsbPackageBuffer buffered) {
                buffered.WriteTo(destination);
            } else {
                destination.Write(((MemoryStream)package).GetBuffer(), 0, checked((int)package.Length));
            }
            if (destination.CanSeek) destination.SetLength(destination.Position);
            return true;
        }

        /// <summary>Emits the same native tabular package from a captured owned worksheet.</summary>
        internal static void WriteOwnedTabular(Stream destination, ExcelDateSystem dateSystem, string sheetName,
            byte[] stylesPart, byte[] themePart, XlsbSharedStringTable? sharedStrings, XlsbDirectTabularPlan plan,
            long maximumSharedStringBytes, CancellationToken cancellationToken) {
            ArraySegment<byte>? sharedStringsPart = sharedStrings?.CreatePart(cancellationToken, maximumSharedStringBytes);
            WriteDirectTabularPackage(dateSystem, destination, sheetName, stylesPart, sharedStringsPart,
                stream => stream.Write(themePart, 0, themePart.Length), plan, cancellationToken);
        }

        private static Action<Stream>? CreateThemeWriter(ThemePart? theme) =>
            theme == null ? null : stream => theme.Theme!.Save(stream);

        private static void WriteDirectTabularPackage(
            ExcelDateSystem dateSystem,
            Stream destination,
            string sheetName,
            byte[]? stylesPart,
            ArraySegment<byte>? sharedStringsPart,
            Action<Stream>? writeTheme,
            XlsbDirectTabularPlan plan,
            CancellationToken cancellationToken) {
            using (ExcelPackageArchive archive = ExcelPackageArchive.Create(destination, ReproducibleEntryTime.DateTime)) {
                WriteDirectTabularWorkbookEntries(archive, dateSystem, sheetName,
                    stylesPart != null, sharedStringsPart.HasValue, writeTheme);
                using (Stream output = archive.OpenEntry("xl/worksheets/sheet1.bin", CompressionLevel.Fastest)) {
                    XlsbWorksheetPartWriter.WriteDirectTabular(output, plan, cancellationToken);
                }
                if (stylesPart != null) archive.WriteEntry("xl/styles.bin", stylesPart, CompressionLevel.Fastest);
                if (sharedStringsPart.HasValue) archive.WriteEntry("xl/sharedStrings.bin", sharedStringsPart.Value, CompressionLevel.Fastest);
            }
        }

        private static void WriteDirectTabularWorkbookEntries(
            ExcelPackageArchive archive,
            ExcelDateSystem dateSystem,
            string sheetName,
            bool hasStyles,
            bool hasSharedStrings,
            Action<Stream>? writeTheme) {
            archive.WriteEntry("[Content_Types].xml", CreateContentTypes(worksheetCount: 1, hasStyles: hasStyles, hasSharedStrings: hasSharedStrings, hasTheme: writeTheme != null), CompressionLevel.Fastest);
            archive.WriteEntry("_rels/.rels", RootRelationships, CompressionLevel.Fastest);
            archive.WriteEntry("xl/workbook.bin", XlsbWorkbookPartWriter.CreateDirectTabular(
                sheetName,
                dateSystem == ExcelDateSystem.NineteenFour), CompressionLevel.Fastest);
            archive.WriteEntry("xl/_rels/workbook.bin.rels", CreateWorkbookRelationships(worksheetCount: 1, hasStyles: hasStyles, hasSharedStrings: hasSharedStrings, hasTheme: writeTheme != null), CompressionLevel.Fastest);
            if (writeTheme == null) return;
            using Stream output = archive.OpenEntry("xl/theme/theme1.xml", CompressionLevel.Fastest);
            writeTheme(output);
        }

        private static void WritePackage(
            ExcelDocument document,
            Stream destination,
            IReadOnlyList<ExcelSheet> sheets,
            byte[]? stylesPart,
            IReadOnlyList<byte[]> worksheetParts,
            IReadOnlyList<XlsbWorksheetHyperlinkPlan> hyperlinkPlans,
            IReadOnlyList<XlsbWorksheetPackageParts> packageParts,
            ArraySegment<byte>? sharedStringsPart,
            XlsbNewWorkbookFormulaContext formulas) {
            ThemePart? theme = document.WorkbookPartRoot.ThemePart;
            XlsbDocumentPropertyParts documentProperties = XlsbDocumentPropertyParts.Capture(document);
            using var positionReportingDestination = destination.CanSeek
                ? null
                : new ExcelPositionReportingWriteStream(destination);
            Stream packageDestination = positionReportingDestination ?? destination;
            using (var archive = new ZipArchive(packageDestination, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "[Content_Types].xml", CreateContentTypes(sheets.Count, stylesPart != null, sharedStringsPart.HasValue, theme != null, documentProperties,
                    packageParts.SelectMany(parts => parts.Parts)));
                WriteEntry(archive, "_rels/.rels", documentProperties.CreateRootRelationships("xl/workbook.bin"));
                documentProperties.WriteEntries((name, content) => WriteEntry(archive, name, content));
                WriteEntry(archive, "xl/workbook.bin", XlsbWorkbookPartWriter.Create(
                    sheets,
                    document.DateSystem == ExcelDateSystem.NineteenFour,
                    XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbWorkbookViewWriter.ElementName),
                    XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbWorkbookProtectionProjector.ElementName),
                    XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbDefinedNameProjector.CollectionName),
                    XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbCalculationPropertiesProjector.ElementName),
                    formulas));
                WriteEntry(archive, "xl/_rels/workbook.bin.rels", CreateWorkbookRelationships(sheets.Count, stylesPart != null, sharedStringsPart.HasValue, theme != null));
                WriteTheme(archive, theme, CompressionLevel.Optimal);
                for (int index = 0; index < worksheetParts.Count; index++) {
                    WriteEntry(archive, WorksheetPartName(index), worksheetParts[index]);
                    if (hyperlinkPlans[index].Relationships.Count != 0 || packageParts[index].Relationships.Count != 0) {
                        WriteEntry(
                            archive,
                            XlsbWorksheetPackageParts.RelationshipPartName(WorksheetPartName(index)),
                            hyperlinkPlans[index].CreateRelationshipPart(packageParts[index].RelationshipsById()));
                    }
                    foreach (XlsbGeneratedPart part in packageParts[index].Parts) {
                        WriteEntry(archive, part.PartName, part.Content);
                        if (part.Relationships.Count != 0) {
                            WriteEntry(archive, XlsbWorksheetPackageParts.RelationshipPartName(part.PartName),
                                XlsbWorksheetPackageParts.CreateRelationshipPart(part.Relationships));
                        }
                    }
                }
                if (stylesPart != null) WriteEntry(archive, "xl/styles.bin", stylesPart);
                if (sharedStringsPart.HasValue) WriteEntry(archive, "xl/sharedStrings.bin", sharedStringsPart.Value);
            }
        }

        // MS-XLSB 2.1.7.52 uses the ordinary DrawingML XML theme part. Preserve
        // it on every writer path; theme-linked style colors must keep their palette.
        private static void WriteTheme(ZipArchive archive, ThemePart? theme, CompressionLevel compression) {
            if (theme == null) return;
            ZipArchiveEntry entry = archive.CreateEntry("xl/theme/theme1.xml", compression);
            entry.LastWriteTime = ReproducibleEntryTime;
            using Stream output = entry.Open();
            theme.Theme!.Save(output);
        }

        private static void ValidateTheme(ExcelDocument document) {
            ThemePart[] themes = document.WorkbookPartRoot.GetPartsOfType<ThemePart>().ToArray();
            if (themes.Length > 1) throw new NotSupportedException("Native XLSB generation supports at most one workbook theme.");
            ThemePart? theme = themes.SingleOrDefault();
            if (theme != null && (theme.Theme == null || theme.Parts.Any()
                || theme.ExternalRelationships.Any() || theme.HyperlinkRelationships.Any())) {
                throw new NotSupportedException("Native XLSB generation requires a standalone DrawingML workbook theme without related parts.");
            }
        }

        private static void ValidateWorkbook(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            if (sheets.Count == 0) {
                throw new NotSupportedException("Native XLSB generation requires at least one worksheet.");
            }

            ThrowIfDuplicateWorkbookElement<WorkbookProperties>(document.WorkbookRoot, "workbook properties");
            ThrowIfDuplicateWorkbookElement<WorkbookProtection>(document.WorkbookRoot, "workbook protection");
            ThrowIfDuplicateWorkbookElement<BookViews>(document.WorkbookRoot, "workbook views");
            ThrowIfDuplicateWorkbookElement<Sheets>(document.WorkbookRoot, "worksheet collections");
            ThrowIfDuplicateWorkbookElement<DefinedNames>(document.WorkbookRoot, "defined-name collections");
            ThrowIfDuplicateWorkbookElement<CalculationProperties>(document.WorkbookRoot, "calculation properties");

            OpenXmlElement? unsupportedWorkbookChild = document.WorkbookRoot.ChildElements
                .FirstOrDefault(element => element is not Sheets
                    && element is not WorkbookProperties
                    && element is not WorkbookProtection
                    && element is not BookViews
                    && element is not DefinedNames
                    && element is not CalculationProperties);
            if (unsupportedWorkbookChild != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook metadata '{unsupportedWorkbookChild.LocalName}'.");
            }

            XlsbWorkbookViewWriter.Validate(XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbWorkbookViewWriter.ElementName), sheets.Count);
            XlsbWorkbookProtectionWriter.Validate(XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbWorkbookProtectionProjector.ElementName));
            XlsbDefinedNameWriter.Validate(document, XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbDefinedNameProjector.CollectionName), sheets);
            XlsbCalculationPropertiesWriter.Validate(XlsbWorkbookMarkup.Child(XlsbWorkbookMarkup.Root(document), XlsbCalculationPropertiesProjector.ElementName));

            WorkbookProperties? properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            if (properties != null) {
                OpenXmlAttribute? unsupported = XlsbOpenXmlAttributeValidator.FindUnsupported(
                    properties, new[] { "date1904" }, IsDefaultWorkbookProperty);
                if (properties.HasChildren || unsupported != null) {
                    throw new NotSupportedException(unsupported == null
                        ? "Native XLSB generation does not yet support workbook property extensions."
                        : $"Native XLSB generation does not yet support workbook property '{unsupported.Value.LocalName}' with value '{unsupported.Value.Value}'.");
                }
            }

            if (document.WorkbookPartRoot.ExternalRelationships.Any()) {
                throw new NotSupportedException("Native XLSB generation does not yet support external workbook relationships.");
            }

            ValidateTheme(document);
            OpenXmlPart? unsupportedPart = document.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .FirstOrDefault(part => part is not WorksheetPart
                    && part is not SharedStringTablePart
                    && part is not WorkbookStylesPart
                    && part is not ThemePart);
            if (unsupportedPart is WorkbookPersonPart) {
                throw new NotSupportedException("Native XLSB generation does not support threaded comments; legacy notes (SetComment) are supported.");
            }
            if (unsupportedPart != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook part '{unsupportedPart.ContentType}'.");
            }
        }

        // The generated BrtWbProp carries date1904. Other attributes are accepted only at the
        // values its zeroed flags already mean, which is how XLS imports state them.
        private static bool IsDefaultWorkbookProperty(OpenXmlAttribute attribute) {
            if (!string.IsNullOrEmpty(attribute.NamespaceUri)) return false;
            switch (attribute.LocalName) {
                case "showObjects":
                    return attribute.Value == "all";
                case "showBorderUnselectedTables":
                case "saveExternalLinkValues":
                    return attribute.Value is "1" or "true";
                case "backupFile":
                case "refreshAllConnections":
                case "filterPrivacy":
                case "hidePivotFieldList":
                case "publishItems":
                case "showInkAnnotation":
                case "promptedSolutions":
                case "checkCompatibility":
                    return attribute.Value is "0" or "false";
                default:
                    return false;
            }
        }

        private static void ValidateDirectTabularWorkbook(
            ExcelDocument document,
            IReadOnlyList<ExcelSheet> sheets,
            string sheetName) {
            if (sheets.Count != 1 || !string.Equals(sheets[0].Name, sheetName, StringComparison.Ordinal)) {
                throw new NotSupportedException("Native XLSB direct tabular generation requires exactly one matching worksheet.");
            }
            if (document.WorkbookRoot.ChildElements.Any(element => element is not Sheets
                && !(element is BookViews views && ExcelDocument.IsNeutralWorkbookViews(views))
                && !(element is WorkbookProperties properties && ExcelDocument.HasOnlyDateSystemWorkbookProperties(properties)))) {
                throw new NotSupportedException("Native XLSB direct tabular generation requires default workbook metadata.");
            }
            if (document.WorkbookPartRoot.ExternalRelationships.Any()) {
                throw new NotSupportedException("Native XLSB generation does not yet support external workbook relationships.");
            }

            ValidateTheme(document);
            OpenXmlPart? unsupportedPart = document.WorkbookPartRoot.Parts
                .Select(pair => pair.OpenXmlPart)
                .FirstOrDefault(part => part is not WorksheetPart
                    && part is not SharedStringTablePart
                    && part is not WorkbookStylesPart
                    && part is not ThemePart);
            if (unsupportedPart != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook part '{unsupportedPart.ContentType}'.");
            }
        }

        private static void ThrowIfDuplicateWorkbookElement<T>(Workbook workbook, string detail)
            where T : OpenXmlElement {
            if (workbook.Elements<T>().Skip(1).Any()) {
                throw new NotSupportedException($"Native XLSB generation does not support multiple {detail} elements.");
            }
        }

        private static string CreateContentTypes(int worksheetCount, bool hasStyles, bool hasSharedStrings, bool hasTheme,
            XlsbDocumentPropertyParts? documentProperties = null, IEnumerable<XlsbGeneratedPart>? generatedParts = null) {
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
            if (hasSharedStrings) {
                builder.Append("<Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>");
            }
            if (hasTheme) builder.Append("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
            documentProperties?.AppendContentTypeOverrides(builder);
            foreach (XlsbGeneratedPart part in generatedParts ?? Enumerable.Empty<XlsbGeneratedPart>()) {
                builder.Append("<Override PartName=\"/");
                builder.Append(System.Security.SecurityElement.Escape(part.PartName));
                builder.Append("\" ContentType=\"");
                builder.Append(System.Security.SecurityElement.Escape(part.ContentType));
                builder.Append("\"/>");
            }
            builder.Append("</Types>");
            return builder.ToString();
        }

        private static string CreateWorkbookRelationships(int worksheetCount, bool hasStyles, bool hasSharedStrings, bool hasTheme) {
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
            if (hasSharedStrings) {
                builder.Append("<Relationship Id=\"rId");
                builder.Append((worksheetCount + (hasStyles ? 2 : 1)).ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.bin\"/>");
            }
            if (hasTheme) {
                builder.Append("<Relationship Id=\"rId");
                builder.Append((worksheetCount + (hasStyles ? 1 : 0) + (hasSharedStrings ? 1 : 0) + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
            }
            builder.Append("</Relationships>");
            return builder.ToString();
        }

        private static void WriteEntry(ZipArchive archive, string name, string content) {
            WriteEntry(archive, name, content, CompressionLevel.Optimal);
        }

        private static void WriteEntry(
            ZipArchive archive,
            string name,
            string content,
            CompressionLevel compressionLevel) {
            WriteEntry(
                archive,
                name,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content),
                compressionLevel);
        }

        private static void WriteEntry(ZipArchive archive, string name, byte[] content) {
            WriteEntry(archive, name, content, CompressionLevel.Optimal);
        }

        private static void WriteEntry(
            ZipArchive archive,
            string name,
            byte[] content,
            CompressionLevel compressionLevel) {
            ZipArchiveEntry entry = archive.CreateEntry(name, compressionLevel);
            entry.LastWriteTime = ReproducibleEntryTime;
            using Stream output = entry.Open();
            output.Write(content, 0, content.Length);
        }

        private static void WriteEntry(ZipArchive archive, string name, ArraySegment<byte> content) {
            WriteEntry(archive, name, content, CompressionLevel.Optimal);
        }

        private static void WriteEntry(
            ZipArchive archive,
            string name,
            ArraySegment<byte> content,
            CompressionLevel compressionLevel) {
            byte[] buffer = content.Array ?? throw new ArgumentException("The package part must have a backing buffer.", nameof(content));
            ZipArchiveEntry entry = archive.CreateEntry(name, compressionLevel);
            entry.LastWriteTime = ReproducibleEntryTime;
            using Stream output = entry.Open();
            output.Write(buffer, content.Offset, content.Count);
        }

        private const string RootRelationships =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "</Relationships>";
    }
}
