using System;
using System.Runtime.CompilerServices;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Core.Internal;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Excel;

internal static partial class ExcelDocumentService
{
    private static readonly ConditionalWeakTable<ExcelDocument, DocumentAssociation> Associations = new();

    private sealed class DocumentAssociation {
        internal string Path { get; set; } = string.Empty;
    }

    public static ExcelDocument CreateFromTemplate(string templatePath, string destinationPath) {
        var document = ExcelDocument.Create(new ExcelCreateOptions { Template = templatePath });
        UpdateSaveAssociation(document, destinationPath);
        return document;
    }

    public static ExcelDocument CreateDocument(string filePath, bool autoSave)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        return ExcelDocument.Create(Path.GetFullPath(filePath), CreateOptions(autoSave));
    }

    public static ExcelDocument CreateInMemoryDocument()
    {
        return ExcelDocument.Create();
    }

    public static ExcelDocument CreateDocumentFromTemplate(string templatePath, string filePath, bool autoSave)
    {
        if (string.IsNullOrWhiteSpace(templatePath))
        {
            throw new ArgumentException("Template path cannot be empty.", nameof(templatePath));
        }

        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        return ExcelDocument.Create(
            Path.GetFullPath(filePath),
            new ExcelCreateOptions
            {
                Template = Path.GetFullPath(templatePath),
                PersistenceMode = autoSave ? DocumentPersistenceMode.SaveOnDispose : DocumentPersistenceMode.Explicit
            });
    }

    public static void CopyWorkbookPackage(string sourcePath, string destinationPath, bool overwrite)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            throw new ArgumentException("Source path cannot be empty.", nameof(sourcePath));
        }

        if (string.IsNullOrWhiteSpace(destinationPath))
        {
            throw new ArgumentException("Destination path cannot be empty.", nameof(destinationPath));
        }

        using var document = ExcelDocument.Load(Path.GetFullPath(sourcePath), new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit
        });
        OfficeFileOutput.Write(Path.GetFullPath(destinationPath), overwrite, temporaryPath => document.SaveCopy(temporaryPath));
    }

    public static ExcelDocument LoadDocument(string filePath, bool readOnly, bool autoSave, string? password = null)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
        }

        if (!string.IsNullOrEmpty(password))
        {
            return OfficeEncryptedPackageService.LoadExcel(resolvedPath, password!, readOnly, autoSave);
        }

        return ExcelDocument.Load(resolvedPath, CreateLoadOptions(readOnly, autoSave));
    }

    public static Task<ExcelDocument> LoadDocumentAsync(
        Uri uri,
        bool readOnly,
        bool allowHttp,
        string? password = null,
        CancellationToken cancellationToken = default)
    {
        if (!string.IsNullOrEmpty(password))
        {
            throw new NotSupportedException("Encrypted remote workbook loads are not supported.");
        }

        var options = ExcelHttpLoadService.CreateOptions(allowHttp);
        options.AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite;
        options.PersistenceMode = DocumentPersistenceMode.Explicit;
        return ExcelDocument.LoadAsync(uri, options, cancellationToken);
    }

    public static void SaveDocument(ExcelDocument document, bool show, string? filePath, string? password = null, ExcelSaveOptions? saveOptions = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        EnsureExplicitPersistence(document);
        string? target = filePath ?? GetAssociatedPath(document);
        if (string.IsNullOrWhiteSpace(target)) throw new InvalidOperationException("No file path provided for save options.");
        target = Path.GetFullPath(target!);
        var options = saveOptions?.Clone() ?? new ExcelSaveOptions();
        if (!string.IsNullOrEmpty(password)) options.Password = password;
        document.Save(target, options);
        UpdateSaveAssociation(document, target);
        CloseDocument(document);
        if (show) FileOpenService.Open(target);
    }

    public static ExcelSaveOptions? CreateSaveOptions(
        bool safePreflight,
        bool safeRepairDefinedNames,
        bool validateOpenXml,
        bool disableFastPackageWriter = false,
        bool evaluateFormulas = false,
        bool clearCachedFormulaResults = false,
        bool markFormulasDirty = false,
        bool forceFullCalculationOnOpen = false)
    {
        if (!safePreflight &&
            !safeRepairDefinedNames &&
            !validateOpenXml &&
            !disableFastPackageWriter &&
            !evaluateFormulas &&
            !clearCachedFormulaResults &&
            !markFormulasDirty &&
            !forceFullCalculationOnOpen)
        {
            return null;
        }

        return new ExcelSaveOptions
        {
            SafePreflight = safePreflight,
            SafeRepairDefinedNames = safeRepairDefinedNames,
            ValidateOpenXml = validateOpenXml,
            DisableFastPackageWriter = disableFastPackageWriter,
            EvaluateFormulasBeforeSave = evaluateFormulas,
            ClearCachedFormulaResultsBeforeSave = clearCachedFormulaResults,
            MarkFormulasDirtyBeforeSave = markFormulasDirty,
            ForceFullCalculationOnOpen = forceFullCalculationOnOpen
        };
    }

    private static ExcelCreateOptions CreateOptions(bool autoSave) => new()
    {
        PersistenceMode = autoSave ? DocumentPersistenceMode.SaveOnDispose : DocumentPersistenceMode.Explicit
    };

    private static ExcelLoadOptions CreateLoadOptions(bool readOnly, bool autoSave) => new()
    {
        AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
        PersistenceMode = autoSave ? DocumentPersistenceMode.SaveOnDispose : DocumentPersistenceMode.Explicit
    };

    public static void CloseDocument(ExcelDocument document)
    {
        if (document == null)
        {
            return;
        }

        try
        {
            document.Dispose();
        }
        finally
        {
            Associations.Remove(document);
        }
    }

    internal static string? GetAssociatedPath(ExcelDocument document) =>
        Associations.TryGetValue(document, out var association) ? association.Path : document.FilePath;

    internal static void UpdateSaveAssociation(ExcelDocument document, string path) {
        var association = Associations.GetValue(document, _ => new DocumentAssociation());
        association.Path = Path.GetFullPath(path);
    }
}
