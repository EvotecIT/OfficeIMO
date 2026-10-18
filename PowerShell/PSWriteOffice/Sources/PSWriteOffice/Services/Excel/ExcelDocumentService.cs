using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Excel;

internal static partial class ExcelDocumentService
{
    private static readonly ConcurrentDictionary<ExcelDocument, string> AssociatedPaths = new();
    private static readonly ConcurrentDictionary<ExcelDocument, string> EncryptedSourcePaths = new();

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

        return ExcelDocument.CreateFromTemplate(
            Path.GetFullPath(templatePath),
            Path.GetFullPath(filePath),
            new ExcelTemplateCreateOptions
            {
                Overwrite = true,
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

        ExcelDocument.CopyPackage(
            Path.GetFullPath(sourcePath),
            Path.GetFullPath(destinationPath),
            overwrite: overwrite);
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
            var document = OfficeEncryptedPackageService.LoadExcel(resolvedPath, password!, readOnly, autoSave);
            AssociatedPaths[document] = resolvedPath;
            EncryptedSourcePaths[document] = resolvedPath;
            return document;
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

        return ExcelDocument.LoadAsync(
            uri,
            ExcelHttpLoadService.CreateOptions(allowHttp),
            CreateLoadOptions(readOnly, autoSave: false),
            cancellationToken);
    }

    public static void SaveDocument(ExcelDocument document, bool show, string? filePath, string? password = null, ExcelSaveOptions? saveOptions = null)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var currentPath = GetAssociatedPath(document) ?? string.Empty;
        if (!string.IsNullOrEmpty(filePath))
        {
            var target = filePath!;
            if (!string.Equals(target, currentPath, StringComparison.OrdinalIgnoreCase))
            {
                SaveDocumentToPath(document, Path.GetFullPath(target), false, password, saveOptions);
                var savedAsPath = Path.GetFullPath(target);
                CloseDocument(document);
                if (show)
                {
                    FileOpenService.Open(savedAsPath);
                }
                return;
            }
        }

        if (!string.IsNullOrEmpty(password))
        {
            var targetPath = GetAssociatedPath(document);
            if (string.IsNullOrWhiteSpace(targetPath))
            {
                throw new InvalidOperationException("No file path provided for encrypted save.");
            }

            OfficeEncryptedPackageService.SaveExcel(document, targetPath!, password!, false, saveOptions);
        }
        else
        {
            if (IsEncryptedSource(document))
            {
                throw new InvalidOperationException("Provide -Password when saving a workbook loaded from an encrypted package.");
            }

            if (saveOptions == null)
            {
                document.Save(currentPath);
            }
            else if (!string.IsNullOrWhiteSpace(currentPath))
            {
                document.Save(currentPath, saveOptions);
            }
            else
            {
                throw new InvalidOperationException("No file path provided for save options.");
            }
        }

        var savedPath = document.FilePath ?? filePath ?? throw new InvalidOperationException("No saved file path was available.");
        CloseDocument(document);
        if (show)
        {
            FileOpenService.Open(savedPath);
        }
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

    private static void SaveDocumentToPath(ExcelDocument document, string path, bool openExcel, string? password, ExcelSaveOptions? saveOptions)
    {
        if (!string.IsNullOrEmpty(password))
        {
            OfficeEncryptedPackageService.SaveExcel(document, path, password!, openExcel, saveOptions);
            return;
        }

        document.Save(path, saveOptions);
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
            AssociatedPaths.TryRemove(document, out _);
            EncryptedSourcePaths.TryRemove(document, out _);
        }
    }

    internal static string? GetAssociatedPath(ExcelDocument document)
    {
        if (AssociatedPaths.TryGetValue(document, out var associatedPath))
        {
            return associatedPath;
        }

        if (!string.IsNullOrWhiteSpace(document.FilePath))
        {
            return document.FilePath;
        }

        return EncryptedSourcePaths.TryGetValue(document, out var path) ? path : null;
    }

    internal static bool IsEncryptedSource(ExcelDocument document) => EncryptedSourcePaths.ContainsKey(document);

    internal static void UpdateSaveAssociation(ExcelDocument document, string path, bool encrypted)
    {
        var resolvedPath = Path.GetFullPath(path);
        AssociatedPaths[document] = resolvedPath;
        if (encrypted)
        {
            EncryptedSourcePaths[document] = resolvedPath;
        }
        else
        {
            EncryptedSourcePaths.TryRemove(document, out _);
        }
    }
}
