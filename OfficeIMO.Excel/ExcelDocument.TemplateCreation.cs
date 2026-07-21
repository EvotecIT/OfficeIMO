using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string MacroEnabledWorkbookContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        private const string TemplateWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
        private const string MacroEnabledTemplateWorkbookContentType = "application/vnd.ms-excel.template.macroEnabled.main+xml";

        /// <summary>
        /// Creates an editable workbook by copying a template package to a destination path and loading the copy.
        /// Package parts such as styles, themes, drawings, tables, named ranges, and worksheet metadata are preserved.
        /// </summary>
        /// <param name="templatePath">Path to the source workbook or template package.</param>
        /// <param name="filePath">Destination workbook path.</param>
        /// <param name="options">Template creation, persistence, and package options.</param>
        /// <returns>The loaded destination workbook.</returns>
        public static ExcelDocument CreateFromTemplate(
            string templatePath,
            string filePath,
            ExcelTemplateCreateOptions? options = null) {
            ExcelTemplateCreateOptions resolved = options ?? new ExcelTemplateCreateOptions();
            CopyPackage(templatePath, filePath, resolved.Overwrite);
            string targetPath = Path.GetFullPath(filePath);
            return Load(targetPath, new ExcelLoadOptions {
                PersistenceMode = resolved.PersistenceMode,
                OpenSettings = resolved.OpenSettings
            });
        }

        /// <summary>
        /// Creates an editable workbook by copying a template package stream to a destination path and loading the copy.
        /// </summary>
        /// <param name="templateStream">Readable template package stream.</param>
        /// <param name="filePath">Destination workbook path.</param>
        /// <param name="options">Template creation, persistence, and package options.</param>
        /// <returns>The loaded destination workbook.</returns>
        public static ExcelDocument CreateFromTemplate(
            Stream templateStream,
            string filePath,
            ExcelTemplateCreateOptions? options = null) {
            ExcelTemplateCreateOptions resolved = options ?? new ExcelTemplateCreateOptions();
            CopyPackage(templateStream, filePath, resolved.Overwrite);
            string targetPath = Path.GetFullPath(filePath);
            return Load(targetPath, new ExcelLoadOptions {
                PersistenceMode = resolved.PersistenceMode,
                OpenSettings = resolved.OpenSettings
            });
        }

        /// <summary>
        /// Copies an existing workbook or template package to a destination path while preserving package parts.
        /// The workbook content type is normalized for the destination extension.
        /// </summary>
        /// <param name="sourcePath">Path to the source workbook or template package.</param>
        /// <param name="destinationPath">Destination workbook path.</param>
        /// <param name="overwrite">When true, replaces an existing destination file.</param>
        public static void CopyPackage(string sourcePath, string destinationPath, bool overwrite = true) {
            if (string.IsNullOrWhiteSpace(sourcePath)) throw new ArgumentNullException(nameof(sourcePath));
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentNullException(nameof(destinationPath));

            string resolvedSourcePath = Path.GetFullPath(sourcePath);
            string resolvedDestinationPath = Path.GetFullPath(destinationPath);
            if (!File.Exists(resolvedSourcePath)) {
                throw new FileNotFoundException($"Workbook package '{resolvedSourcePath}' was not found.", resolvedSourcePath);
            }

            EnsureMacroCompatibleCopy(resolvedSourcePath, resolvedDestinationPath);
            EnsureDestinationDirectory(resolvedDestinationPath);
            File.Copy(resolvedSourcePath, resolvedDestinationPath, overwrite);
            NormalizeTemplateWorkbookContentType(resolvedDestinationPath);
        }

        /// <summary>
        /// Copies a workbook or template package stream to a destination path while preserving package parts.
        /// The workbook content type is normalized for the destination extension.
        /// </summary>
        /// <param name="sourceStream">Readable workbook package stream.</param>
        /// <param name="destinationPath">Destination workbook path.</param>
        /// <param name="overwrite">When true, replaces an existing destination file.</param>
        public static void CopyPackage(Stream sourceStream, string destinationPath, bool overwrite = true) {
            if (sourceStream == null) throw new ArgumentNullException(nameof(sourceStream));
            if (!sourceStream.CanRead) throw new ArgumentException("Workbook package stream must be readable.", nameof(sourceStream));
            if (string.IsNullOrWhiteSpace(destinationPath)) throw new ArgumentNullException(nameof(destinationPath));

            string targetPath = Path.GetFullPath(destinationPath);
            OfficeFileCommit.WriteAllBytes(
                targetPath,
                OfficeStreamReader.ReadAllBytes(sourceStream),
                overwrite ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);

            try {
                EnsureMacroCompatibleCopy(targetPath, targetPath);
                NormalizeTemplateWorkbookContentType(targetPath);
            } catch {
                File.Delete(targetPath);
                throw;
            }
        }

        private static void EnsureDestinationDirectory(string filePath) {
            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                Directory.CreateDirectory(directory);
            }
        }

        private static void NormalizeTemplateWorkbookContentType(string filePath) {
            string? contentType = ResolveWorkbookContentTypeForPath(filePath);
            if (contentType == null) {
                return;
            }

            using ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml");
            if (contentTypesEntry == null) {
                return;
            }

            XDocument document;
            using (Stream stream = contentTypesEntry.Open()) {
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }

            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement? workbookOverride = document
                .Root?
                .Elements(ns + "Override")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute("PartName"), "/xl/workbook.xml", StringComparison.OrdinalIgnoreCase));
            if (workbookOverride == null) {
                return;
            }

            XAttribute? contentTypeAttribute = workbookOverride.Attribute("ContentType");
            if (string.Equals(contentTypeAttribute?.Value, contentType, StringComparison.Ordinal)) {
                return;
            }

            workbookOverride.SetAttributeValue("ContentType", contentType);
            contentTypesEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using Stream replacementStream = replacement.Open();
            document.Save(replacementStream, SaveOptions.DisableFormatting);
        }

        private static void EnsureMacroCompatibleCopy(string sourcePath, string destinationPath) {
            if (AllowsMacroWorkbookContent(destinationPath) || !ForcesMacroFreeWorkbookContent(destinationPath)) {
                return;
            }

            if (PackageContainsMacroContent(sourcePath)) {
                throw new InvalidOperationException("Macro-enabled workbook packages cannot be copied to a macro-free .xlsx or .xltx destination. Use .xlsm or .xltm to preserve macros.");
            }
        }

        private static bool PackageContainsMacroContent(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            if (archive.GetEntry("xl/vbaProject.bin") != null) {
                return true;
            }

            string? workbookContentType = ReadWorkbookOverrideContentType(archive);
            return string.Equals(workbookContentType, MacroEnabledWorkbookContentType, StringComparison.Ordinal)
                || string.Equals(workbookContentType, MacroEnabledTemplateWorkbookContentType, StringComparison.Ordinal);
        }

        private static string? ReadWorkbookOverrideContentType(ZipArchive archive) {
            ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml");
            if (contentTypesEntry == null) {
                return null;
            }

            using Stream stream = contentTypesEntry.Open();
            XDocument document = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
            return document
                .Root?
                .Elements(ns + "Override")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute("PartName"), "/xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
                ?.Attribute("ContentType")
                ?.Value;
        }

        private static bool AllowsMacroWorkbookContent(string filePath) {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension is ".xlsm" or ".xltm";
        }

        private static bool ForcesMacroFreeWorkbookContent(string filePath) {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension is ".xlsx" or ".xltx";
        }

        private static string? ResolveWorkbookContentTypeForPath(string filePath) {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch {
                ".xlsx" => WorkbookContentType,
                ".xlsm" => MacroEnabledWorkbookContentType,
                ".xltx" => TemplateWorkbookContentType,
                ".xltm" => MacroEnabledTemplateWorkbookContentType,
                ".xlam" => ExcelPackageUtilities.AddInWorkbookContentType,
                _ => null
            };
        }
    }
}
