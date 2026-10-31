using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

internal enum OfficeMergeFamily { Word, Excel, Pdf }

/// <summary>Owns only shell-loaded input lifetimes and dispatches to native append/merge methods.</summary>
internal static class OfficeMergeDispatch {
    internal static OfficeMergeFamily GetFamily(string extension) {
        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return OfficeMergeFamily.Pdf;
        if (WordFormatCatalog.TryGetByExtension(extension, out var word) && word!.Encoding == OfficeFormatEncoding.OpenXml)
            return OfficeMergeFamily.Word;
        if (ExcelFormatCatalog.TryGetByExtension(extension, out var excel) && excel!.Encoding == OfficeFormatEncoding.OpenXml)
            return OfficeMergeFamily.Excel;
        throw new NotSupportedException("Native merge supports PDF and Open XML Word/Excel packages only. Unsupported format: " + extension);
    }

    /// <summary>Checks the live target without saving, changing persistence, or taking ownership.</summary>
    internal static string? ValidateDocument(object document, string extension) {
        switch (document) {
            case WordDocument word:
                if (word.AccessMode != DocumentAccessMode.ReadWrite) throw new PSArgumentException("Document must be editable.");
                ValidateFormat(word.SourceFormatDescriptor, extension);
                if (word.InspectSignatures().HasSignatures)
                    throw new PSArgumentException("Signed Word merge targets cannot be changed without invalidating their signatures.");
                return word.FilePath;
            case ExcelDocument excel:
                if (excel.AccessMode != DocumentAccessMode.ReadWrite) throw new PSArgumentException("Document must be editable.");
                ValidateFormat(excel.SourceFormatDescriptor, extension);
                if (excel.InspectSignatures().HasSignatures)
                    throw new PSArgumentException("Signed Excel merge targets cannot be changed without invalidating their signatures.");
                return excel.FilePath;
            case PdfDocument:
                throw new NotSupportedException("Live PDF targets cannot be appended in place. Use Path or LiteralPath and OutputPath for a native PDF merge.");
            default:
                throw new NotSupportedException("Document must be an editable WordDocument or ExcelDocument. Other live formats do not support this native append contract.");
        }
    }

    internal static void Execute(IReadOnlyList<string> paths, SecureString[]? passwords, OfficeMergeFamily family,
        string extension, object options, object? document, string? output, bool noClobber) {
        var owned = new List<IDisposable>();
        Exception? operationFailure = null;
        try {
            // Open every source before touching a caller-owned target. A later bad password/file cannot partially append earlier inputs.
            for (int index = 0; index < paths.Count; index++) {
                SecureString? secure = passwords == null ? null : passwords[passwords.Length == 1 ? 0 : index];
                string? password = secure == null ? null : new NetworkCredential(string.Empty, secure).Password;
                bool primary = document == null && index == 0;
                IDisposable source = family switch {
                    OfficeMergeFamily.Word => WordDocument.Load(paths[index], new WordLoadOptions {
                        AccessMode = primary ? DocumentAccessMode.ReadWrite : DocumentAccessMode.ReadOnly,
                        PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false, Password = password
                    }),
                    OfficeMergeFamily.Excel => ExcelDocument.Load(paths[index], new ExcelLoadOptions {
                        AccessMode = primary ? DocumentAccessMode.ReadWrite : DocumentAccessMode.ReadOnly,
                        PersistenceMode = DocumentPersistenceMode.Explicit, AllowBinaryProjection = false, Password = password
                    }),
                    _ => PdfDocument.Load(paths[index], new PdfLoadOptions { Password = password })
                };
                owned.Add(source);
                // Extension checking alone must not permit a mislabeled package or a template conversion.
                if (source is WordDocument word) ValidateFormat(word.SourceFormatDescriptor, extension);
                if (source is ExcelDocument excel) ValidateFormat(excel.SourceFormatDescriptor, extension);
                if (primary && family != OfficeMergeFamily.Pdf) ValidateDocument(source, extension);
            }
            if (family == OfficeMergeFamily.Pdf) {
                // Keep each loaded source's native read options and permission enforcement intact.
                using var merged = PdfDocument.Merge(owned.Cast<PdfDocument>(), (PdfMergeOptions)options);
                OfficeOutput.Write(output!, stream => merged.Save(stream), noClobber);
                return;
            }
            object target = document ?? owned[0];
            int firstAppend = document == null ? 1 : 0;
            if (target is WordDocument targetWord) {
                for (int index = firstAppend; index < owned.Count; index++)
                    targetWord.AppendDocument((WordDocument)owned[index], (WordDocumentMergeOptions)options);
                if (output != null) OfficeOutput.Write(output, stream => targetWord.Save(stream), noClobber);
            } else {
                var targetExcel = (ExcelDocument)target;
                var settings = (ExcelWorkbookMergeOptions)options;
                // Resolve selected sheets across all appended workbooks before mutating the target.
                if (settings.SheetNames != null)
                    for (int index = firstAppend; index < owned.Count; index++)
                        foreach (string name in settings.SheetNames) ((ExcelDocument)owned[index]).GetSheet(name);
                for (int index = firstAppend; index < owned.Count; index++)
                    targetExcel.MergeWorkbookFrom((ExcelDocument)owned[index], settings);
                if (output != null) OfficeOutput.Write(output, stream => targetExcel.Save(stream), noClobber);
            }
        } catch (Exception error) {
            operationFailure = error;
            throw;
        } finally {
            // Includes the partially loaded prefix when a later Load or Merge throws; the live target is never in this list.
            var disposalFailures = new List<Exception>();
            foreach (IDisposable input in owned) {
                try { input.Dispose(); }
                catch (Exception error) { disposalFailures.Add(error); }
            }
            if (operationFailure == null && disposalFailures.Count > 0) throw new AggregateException(disposalFailures);
        }
    }

    private static void ValidateFormat(OfficeFormatDescriptor format, string extension) {
        if (format.Encoding != OfficeFormatEncoding.OpenXml || !format.Extension.Equals(extension, StringComparison.OrdinalIgnoreCase))
            throw new PSArgumentException("The native document format must match every input and OutputPath; format conversion is not part of merge.");
    }
}
