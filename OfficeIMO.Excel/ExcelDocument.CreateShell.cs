using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// A new workbook starts as a package OfficeIMO writes itself: the workbook part, the default theme and stylesheet,
    /// and the application properties, as a creation through the model writes them. It then works like a loaded workbook
    /// that has not opened its SDK package; the bridge opens the SDK over these bytes when something needs the model.
    /// </summary>
    public partial class ExcelDocument {
        private const string CreatedWorkbookPartName = "/xl/workbook.xml";
        private const string CreatedWorkbookXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:workbook xmlns:x=\"" + SpreadsheetMainNamespace + "\" />";
        // The model's default stylesheet; font zero, the Normal font, takes the creation theme's body size and family.
        private const string CreatedStylesXmlStart = "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:styleSheet xmlns:x=\"" + SpreadsheetMainNamespace + "\">"
            + "<x:fonts count=\"1\"><x:font><x:sz val=\"";
        private const string CreatedStylesXmlEnd = "\" /></x:font></x:fonts>"
            + "<x:fills count=\"2\"><x:fill><x:patternFill patternType=\"none\" /></x:fill><x:fill><x:patternFill patternType=\"gray125\" /></x:fill></x:fills>"
            + "<x:borders count=\"1\"><x:border /></x:borders>"
            + "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>"
            + "<x:cellXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" /></x:cellXfs>"
            + "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>"
            + "<x:dxfs count=\"0\" /><x:tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\" /></x:styleSheet>";
        private const string CreatedApplicationPropertiesXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\""
            + " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>OfficeIMO.Excel</Application></Properties>";

        // A sheet the model adds to a new workbook gets the default selection when the save tidies its views. The
        // unprefixed form is the one the direct writers produce, which the readers' fast paths recognise.
        private const string CreatedWorksheetXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns=\"" + SpreadsheetMainNamespace
            + "\"><sheetViews><sheetView workbookViewId=\"0\"><selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews><sheetData/></worksheet>";

        /// <summary>Creates a detached workbook that must be saved explicitly to a path or stream.</summary>
        /// <param name="options">Creation options. SaveOnDispose is invalid without an associated destination.</param>
        public static ExcelDocument Create(ExcelCreateOptions? options = null) {
            ExcelCreateOptions resolved = options ?? new ExcelCreateOptions();
            if (resolved.Template != null || resolved.TemplateStream != null) return CreateFromUnifiedOptions(null, null, resolved);
            if (resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }
            return CreateOwnedDocument(resolved, resolved.DocumentType, filePath: null, sourceStream: null);
        }

        /// <summary>
        /// Creates a new Excel document at the specified path.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <param name="options">Creation and persistence options.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath, ExcelCreateOptions? options = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            ExcelCreateOptions resolved = options ?? new ExcelCreateOptions();
            if (resolved.Template != null || resolved.TemplateStream != null) return CreateFromUnifiedOptions(filePath, null, resolved);
            return CreateOwnedDocument(resolved, CreatedDocumentTypeOf(filePath) ?? resolved.DocumentType, filePath, sourceStream: null);
        }

        /// <summary>
        /// Creates a new Excel document in memory and optionally persists it to the provided stream on dispose.
        /// </summary>
        /// <param name="stream">Destination stream to receive the workbook package.</param>
        /// <param name="options">Creation and persistence options.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(Stream stream, ExcelCreateOptions? options = null) {
            OfficeDocumentLifecycle.EnsureAssociatedDestination(stream, nameof(stream));
            ExcelCreateOptions resolved = options ?? new ExcelCreateOptions();
            if (resolved.Template != null || resolved.TemplateStream != null) return CreateFromUnifiedOptions(null, stream, resolved);
            return CreateOwnedDocument(resolved, resolved.DocumentType, filePath: null, stream);
        }

        // Set for a workbook created through the owned package, so the model treats it as new when the SDK opens.
        private bool _deferredCreated;

        /// <summary>The main content type of a new workbook of <paramref name="type"/>, or null for one the model creates.</summary>
        private static string? CreatedMainContentType(ExcelDocumentType type) => type switch {
            ExcelDocumentType.Workbook => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            ExcelDocumentType.Template => "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
            ExcelDocumentType.MacroEnabledWorkbook => "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            ExcelDocumentType.MacroEnabledTemplate => "application/vnd.ms-excel.template.macroEnabled.main+xml",
            ExcelDocumentType.AddIn => "application/vnd.ms-excel.addin.macroEnabled.main+xml",
            _ => null
        };

        /// <summary>The document type a destination's extension asks for, or null when it names none.</summary>
        private static ExcelDocumentType? CreatedDocumentTypeOf(string? filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) return null;
            return Path.GetExtension(filePath).ToLowerInvariant() switch {
                ".xlsx" => ExcelDocumentType.Workbook,
                ".xlsm" => ExcelDocumentType.MacroEnabledWorkbook,
                ".xltx" => ExcelDocumentType.Template,
                ".xltm" => ExcelDocumentType.MacroEnabledTemplate,
                ".xlam" => ExcelDocumentType.AddIn,
                _ => null
            };
        }

        /// <summary>
        /// Adds the application properties a save through the model gives a new workbook, unless the package has them or
        /// another part holds their usual name.
        /// </summary>
        private static void CreateApplicationPropertiesIfMissing(OfficeOpcEditablePackage package, bool related) {
            const string partName = "/docProps/app.xml";
            if (related || package.PartExists(partName)) return;
            package.CreateBinaryPart(partName, "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                Encoding.UTF8.GetBytes(CreatedApplicationPropertiesXml));
            package.AddRelationship(null, ApplicationPropertiesRelationshipType, "docProps/app.xml");
        }

        /// <summary>
        /// Points the workbook's sheet entries at the parts of a package another writer produced, matching them by name;
        /// false when the package's sheets differ.
        /// </summary>
        private bool TryRebindDeferredSheets(byte[] bytes) {
            ExcelDeferredWorkbook? fresh = TryReadDeferredWorkbook(bytes);
            ExcelDeferredWorkbook current = _deferredWorkbook!;
            if (fresh == null || fresh.Sheets.Count != current.Sheets.Count
                || !string.Equals(fresh.WorkbookPartName, current.WorkbookPartName, StringComparison.Ordinal)) {
                return false;
            }
            for (int index = 0; index < current.Sheets.Count; index++) {
                if (!string.Equals(fresh.Sheets[index].Name, current.Sheets[index].Name, StringComparison.Ordinal)) return false;
            }
            for (int index = 0; index < current.Sheets.Count; index++) {
                current.Sheets[index].RelationshipId = fresh.Sheets[index].RelationshipId;
                current.Sheets[index].PartName = fresh.Sheets[index].PartName;
            }
            return true;
        }

        /// <summary>Writes the package of a new, empty workbook.</summary>
        private static string CreatedStylesXml(OfficeTheme? theme) {
            OfficeTheme source = theme ?? OfficeTheme.Default;
            return CreatedStylesXmlStart + source.TypeScale.Body.ToPoints().ToString("R", System.Globalization.CultureInfo.InvariantCulture)
                + "\" /><x:name val=\"" + ExcelPartEdits.Escape(source.Fonts.Body.Family) + CreatedStylesXmlEnd;
        }

        private static byte[] CreateWorkbookShell(string mainContentType, OfficeTheme? theme) {
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Create();
            package.CreateBinaryPart(CreatedWorkbookPartName, mainContentType, Encoding.UTF8.GetBytes(CreatedWorkbookXml));
            package.AddRelationship(null, OfficeOpcPackage.OfficeDocumentRelationshipType, "xl/workbook.xml");
            byte[] themeBytes = theme == null ? DefaultThemeBytes.Value : Encoding.UTF8.GetBytes(CreateWorkbookThemeXml(theme));
            package.CreateBinaryPart("/xl/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml", themeBytes);
            package.AddRelationship(CreatedWorkbookPartName, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "theme", "theme/theme1.xml");
            package.CreateBinaryPart("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                Encoding.UTF8.GetBytes(CreatedStylesXml(theme)));
            package.AddRelationship(CreatedWorkbookPartName, StylesRelationshipType, "styles.xml");
            using var output = new MemoryStream(16 * 1024);
            package.Save(output);
            return output.ToArray();
        }
    }
}
