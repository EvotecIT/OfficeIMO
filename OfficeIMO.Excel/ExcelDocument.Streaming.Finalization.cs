using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using E = DocumentFormat.OpenXml.ExtendedProperties;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private static bool NeedsStreamingFinalization(ExcelWriterOptions options) =>
        options.DocumentProperties?.Count > 0 || options.SaveOptions != null;

    private static void FinalizeStreamingFile(string path, ExcelWriterOptions options) {
        ExcelSaveOptions? save = options.SaveOptions;
        CheckStreamingTemporaryBudget(path, save);
        if (options.DocumentProperties?.Count > 0) {
            // Open only package metadata; worksheet DOMs are never requested by this path.
            using var package = SpreadsheetDocument.Open(path, true);
            foreach (var property in options.DocumentProperties) SetStreamingDocumentProperty(package, property.Key, property.Value);
        }
        if (save == null || !RequiresStreamingSavePass(save)) {
            CheckStreamingTemporaryBudget(path, save);
            return;
        }
        if (save.MaxInMemoryPackageBytes is long memoryLimit && new FileInfo(path).Length > memoryLimit)
            throw new InvalidOperationException("The requested native save policy exceeds MaxInMemoryPackageBytes.");

        // These explicitly requested policies use the existing save owner, including its validation,
        // formula, encryption and format rules. Default row ingestion never enters this DOM phase.
        string finalPath = OfficeFileCommit.CreateTemporaryPath(path);
        try {
            using (var document = Load(path))
            using (var destination = new FileStream(finalPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                document.Save(destination, save);
            }
            CheckStreamingTemporaryBudget(finalPath, save);
            OfficeFileCommit.CommitTemporaryFileAtomically(finalPath, path);
        } finally {
            if (File.Exists(finalPath)) File.Delete(finalPath);
        }
    }

    private static void CheckStreamingTemporaryBudget(string path, ExcelSaveOptions? options) {
        if (options?.MaxTemporaryPackageBytes is long maximum && new FileInfo(path).Length > maximum)
            throw new InvalidOperationException("The completed package exceeds MaxTemporaryPackageBytes.");
    }

    private static bool RequiresStreamingSavePass(ExcelSaveOptions options) =>
        options.Format != ExcelFileFormat.Xlsx || options.ResolvePassword(null) != null || options.ValidateOpenXml
        || options.SafePreflight || options.SafeRepairDefinedNames || options.DisableFastPackageWriter
        || options.EvaluateFormulasBeforeSave || options.ClearCachedFormulaResultsBeforeSave
        || options.MarkFormulasDirtyBeforeSave || options.ForceFullCalculationOnOpen;

    private static void ValidateStreamingDocumentProperties(IReadOnlyDictionary<string, object?>? properties) {
        if (properties == null) return;
        foreach (string name in properties.Keys) {
            switch (name.ToUpperInvariant()) {
                case "TITLE": case "SUBJECT": case "CREATOR": case "AUTHOR": case "KEYWORDS":
                case "DESCRIPTION": case "CATEGORY": case "REVISION": case "LASTMODIFIEDBY": case "VERSION":
                case "CREATED": case "MODIFIED": case "LASTPRINTED": case "COMPANY": case "MANAGER":
                case "APPLICATION": case "APPLICATIONNAME": break;
                default: throw new ArgumentException($"'{name}' is not a supported Excel document property.", nameof(properties));
            }
        }
    }

    private static void SetStreamingDocumentProperty(SpreadsheetDocument package, string name, object? value) {
        string? text = value == null ? null : System.Convert.ToString(value, CultureInfo.InvariantCulture);
        DateTime? Date() => value == null ? null : value is DateTimeOffset offset ? offset.DateTime : System.Convert.ToDateTime(value, CultureInfo.InvariantCulture);
        switch (name.ToUpperInvariant()) {
            case "TITLE": package.PackageProperties.Title = text; return;
            case "SUBJECT": package.PackageProperties.Subject = text; return;
            case "AUTHOR": case "CREATOR": package.PackageProperties.Creator = text; return;
            case "KEYWORDS": package.PackageProperties.Keywords = text; return;
            case "DESCRIPTION": package.PackageProperties.Description = text; return;
            case "CATEGORY": package.PackageProperties.Category = text; return;
            case "REVISION": package.PackageProperties.Revision = text; return;
            case "LASTMODIFIEDBY": package.PackageProperties.LastModifiedBy = text; return;
            case "VERSION": package.PackageProperties.Version = text; return;
            case "CREATED": package.PackageProperties.Created = Date(); return;
            case "MODIFIED": package.PackageProperties.Modified = Date(); return;
            case "LASTPRINTED": package.PackageProperties.LastPrinted = Date(); return;
        }
        var part = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
        var properties = part.Properties ??= new E.Properties();
        switch (name.ToUpperInvariant()) {
            case "COMPANY": properties.Company = new E.Company(text ?? string.Empty); break;
            case "MANAGER": properties.Manager = new E.Manager(text ?? string.Empty); break;
            case "APPLICATION": case "APPLICATIONNAME": properties.Application = new E.Application(text ?? string.Empty); break;
            default: throw new ArgumentException($"'{name}' is not a supported Excel document property.", nameof(name));
        }
    }
}
