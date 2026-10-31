using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

internal enum OfficeProtectionFamily { Word, Excel, PowerPoint, Pdf }

/// <summary>Dispatches protection to native owners; never takes ownership of live input.</summary>
internal static class OfficeProtectionDispatch {
    internal static OfficeProtectionFamily Family(string extension) {
        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return OfficeProtectionFamily.Pdf;
        if (WordFormatCatalog.TryGetByExtension(extension, out var word) && word!.Encoding == OfficeFormatEncoding.OpenXml)
            return OfficeProtectionFamily.Word;
        if (ExcelFormatCatalog.TryGetByExtension(extension, out var excel) && excel!.Encoding == OfficeFormatEncoding.OpenXml)
            return OfficeProtectionFamily.Excel;
        if (PowerPointFormatCatalog.TryGetByExtension(extension, out var presentation) && presentation!.Encoding == OfficeFormatEncoding.OpenXml)
            return OfficeProtectionFamily.PowerPoint;
        throw new NotSupportedException("Protection supports PDF and native Open XML Word, Excel and PowerPoint packages. Unsupported format: " + extension);
    }

    internal static IDisposable Load(string path, OfficeProtectionFamily family) {
        // Password is the OUTPUT password. Never reuse it as input authentication or project a binary source.
        try {
            return family switch {
                OfficeProtectionFamily.Word => WordDocument.Load(path, new WordLoadOptions {
                    AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false
                }),
                OfficeProtectionFamily.Excel => ExcelDocument.Load(path, new ExcelLoadOptions {
                    AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit, AllowBinaryProjection = false
                }),
                OfficeProtectionFamily.PowerPoint => PowerPointDocument.Load(path, new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadWrite, PersistenceMode = DocumentPersistenceMode.Explicit, AllowLegacyProjection = false
                }),
                _ => PdfDocument.Load(path)
            };
        } catch (PdfInvalidPasswordException error) {
            throw new PSArgumentException("Encrypted PDF input requires an input-password path unavailable on this command. The source was not changed.", error);
        } catch (InvalidDataException error) when (error.Message.IndexOf("encrypt", StringComparison.OrdinalIgnoreCase) >= 0 ||
            error.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0) {
            throw new PSArgumentException("Encrypted input requires an input-password path unavailable on this command. The source was not changed.", error);
        }
    }

    internal static string Extension(object document) => document switch {
        WordDocument word => word.SourceFormatDescriptor.Extension,
        ExcelDocument excel => excel.SourceFormatDescriptor.Extension,
        PowerPointDocument presentation => presentation.SourceFormatDescriptor.Extension,
        PdfDocument => ".pdf",
        _ => throw new PSArgumentException("Document must be a native WordDocument, ExcelDocument, PowerPointPresentation or PdfDocument.")
    };

    internal static string? AssociatedPath(object document) => document switch {
        WordDocument word => word.FilePath,
        ExcelDocument excel => excel.FilePath,
        PowerPointDocument presentation => presentation.FilePath,
        PdfDocument pdf => pdf.FilePath,
        _ => null
    };

    internal static void Validate(object document, string extension, bool changes, bool encrypt, string? output) {
        string actual = Extension(document);
        if (!actual.Equals(extension, StringComparison.OrdinalIgnoreCase))
            throw new PSArgumentException("The filename format must match the native document format; protection does not convert formats.");
        switch (document) {
            case WordDocument word:
                CheckEditable(word.AccessMode, changes);
                if (word.SourceFormatDescriptor.Encoding != OfficeFormatEncoding.OpenXml) RejectBinary();
                if (changes && word.InspectSignatures().HasSignatures) RejectSigned();
                // A stream export intentionally does not replace Word's associated password. SaveOnDispose could
                // overwrite an in-place encrypted export with plaintext; require explicit persistence for that case.
                if (encrypt && word.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && SamePath(word.FilePath, output))
                    throw new PSArgumentException("In-place live Word encryption requires Explicit persistence. Use an independent OutputPath or load with PersistenceMode Explicit.");
                break;
            case ExcelDocument excel:
                CheckEditable(excel.AccessMode, changes);
                if (excel.SourceFormatDescriptor.Encoding != OfficeFormatEncoding.OpenXml) RejectBinary();
                if (changes && excel.InspectSignatures().HasSignatures) RejectSigned();
                if (encrypt && excel.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && SamePath(excel.FilePath, output))
                    throw new PSArgumentException("In-place live Excel encryption requires Explicit persistence. Use an independent OutputPath or load with PersistenceMode Explicit.");
                break;
            case PowerPointDocument presentation:
                CheckEditable(presentation.AccessMode, changes);
                if (presentation.SourceFormatDescriptor.Encoding != OfficeFormatEncoding.OpenXml) RejectBinary();
                if (changes && presentation.InspectSignatures().HasSignatureMetadata) RejectSigned();
                if (encrypt && presentation.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && SamePath(presentation.FilePath, output))
                    throw new PSArgumentException("In-place live PowerPoint encryption requires Explicit persistence. Use an independent OutputPath or load with PersistenceMode Explicit.");
                break;
            case PdfDocument pdf:
                if (!changes) break;
                PdfDocumentSecurityInfo security;
                try { security = pdf.InspectGeometryForViewing().Security; }
                catch (PdfInvalidPasswordException error) {
                    throw new PSArgumentException("Encrypted PDF input requires an input-password path unavailable on Protect-OfficeDocument.", error);
                }
                if (security.HasEncryption)
                    throw new PSArgumentException("Encrypted PDF input requires current owner authentication and Reencrypt; this input-password path is unavailable on Protect-OfficeDocument.");
                if (security.HasSignatures) RejectSigned();
                break;
        }
    }

    internal static bool Supports(OfficeProtectionFamily family, OfficeProtectionRestriction restriction) => family switch {
        OfficeProtectionFamily.Word => restriction is OfficeProtectionRestriction.ReadOnly or OfficeProtectionRestriction.Comments or
            OfficeProtectionRestriction.TrackedChanges or OfficeProtectionRestriction.Forms,
        OfficeProtectionFamily.Excel => restriction is OfficeProtectionRestriction.Structure or OfficeProtectionRestriction.Windows or OfficeProtectionRestriction.ReadOnlyRecommended,
        OfficeProtectionFamily.Pdf => restriction is OfficeProtectionRestriction.Print or OfficeProtectionRestriction.Copy or OfficeProtectionRestriction.Modify or
            OfficeProtectionRestriction.Annotate or OfficeProtectionRestriction.FillForms or OfficeProtectionRestriction.Assemble,
        _ => false
    };

    /// <summary>Applies already validated settings and serializes before creating any destination directory.</summary>
    internal static object Apply(object document, OfficeProtectionRestriction[] restrictions, OfficeProtectionOptionBinding options,
        bool encrypt, string? password, string? restrictionPassword, string? output, Action<string> verbose) {
        switch (document) {
            case WordDocument word:
                if (restrictions.Length > 0) {
                    word.Settings.ProtectionPassword = restrictionPassword!;
                    word.Settings.SetProtectionType((WordDocumentProtectionType)Enum.Parse(typeof(WordDocumentProtectionType), restrictions[0].ToString()));
                }
                if (output != null) {
                    var save = new WordSaveOptions();
                    if (encrypt) save.Password = password; // Unset preserves encryption on an authenticated live source.
                    Publish(output, stream => word.Save(stream, save));
                }
                return word;
            case ExcelDocument excel:
                if (options.HasWorkbook) excel.ProtectWorkbook(options.Workbook);
                if (options.HasReservation) excel.SetWriteReservation(options.Reservation);
                if (output != null) {
                    var save = new ExcelSaveOptions();
                    if (encrypt) save.Password = password;
                    Publish(output, stream => excel.Save(stream, save));
                }
                return excel;
            case PowerPointDocument presentation:
                var presentationSave = new PowerPointSaveOptions();
                if (encrypt) presentationSave.Password = password;
                PublishPresentation(presentation, output!, presentationSave);
                return presentation;
            case PdfDocument pdf:
                var encryption = options.Pdf!;
                foreach (var restriction in restrictions) encryption.AllowedPermissions &= ~Permission(restriction);
                // The native owner requires a full rewrite and proves non-security structure preservation.
                var result = pdf.Security.Encrypt(encryption);
                result.PreservationReport.ThrowIfFailed();
                verbose(result.PreservationReport.Summary);
                if (output != null) {
                    byte[] bytes = result.Pdf;
                    OfficeOutput.Write(output, stream => stream.Write(bytes, 0, bytes.Length), noClobber: false);
                }
                // Allocate only after successful publication; the command owns this replacement until PassThru succeeds.
                return result.ToDocument();
            default:
                throw new PSArgumentException("Unsupported native Document.");
        }
    }

    private static PdfStandardPermissions Permission(OfficeProtectionRestriction restriction) => restriction switch {
        OfficeProtectionRestriction.Print => PdfStandardPermissions.Print | PdfStandardPermissions.HighQualityPrint,
        OfficeProtectionRestriction.Copy => PdfStandardPermissions.CopyContents,
        OfficeProtectionRestriction.Modify => PdfStandardPermissions.ModifyContents,
        OfficeProtectionRestriction.Annotate => PdfStandardPermissions.ModifyAnnotations,
        OfficeProtectionRestriction.FillForms => PdfStandardPermissions.FillForms,
        OfficeProtectionRestriction.Assemble => PdfStandardPermissions.AssembleDocument,
        _ => PdfStandardPermissions.None
    };

    private static void Publish(string output, Action<Stream> save) {
        using var serialized = new MemoryStream();
        save(serialized); // Native signature, encryption and preservation policies run before publication.
        serialized.Position = 0;
        OfficeOutput.Write(output, stream => serialized.CopyTo(stream), noClobber: false);
    }

    private static void PublishPresentation(PowerPointDocument presentation, string output, PowerPointSaveOptions options) {
        // PowerPoint's stream format selector cannot represent templates, slide shows or add-ins. The native
        // SaveCopy path preserves their package kind and VBA without changing the caller's associated path.
        string directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-protection-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string staged = System.IO.Path.Combine(directory, "protected" + System.IO.Path.GetExtension(output));
            presentation.SaveCopy(staged, options);
            using var source = File.OpenRead(staged);
            OfficeOutput.Write(output, stream => source.CopyTo(stream), noClobber: false);
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static bool SamePath(string? source, string? output) => !string.IsNullOrWhiteSpace(source) && output != null &&
        OfficePath.AreEquivalent(source!, output);
    private static void CheckEditable(DocumentAccessMode access, bool changes) {
        if (changes && access != DocumentAccessMode.ReadWrite) throw new PSArgumentException("Document must be editable.");
    }
    private static void RejectBinary() => throw new PSArgumentException("Only native Open XML packages are supported; binary projection is not protection.");
    private static void RejectSigned() => throw new PSArgumentException("Native safe-save policy blocks protection of signed documents because rewriting can invalidate signatures.");
}
