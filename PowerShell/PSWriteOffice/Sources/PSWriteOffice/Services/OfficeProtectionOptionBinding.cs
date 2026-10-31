using System;
using System.Collections;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Captures native protection settings without modifying caller-owned option objects.</summary>
internal sealed class OfficeProtectionOptionBinding {
    internal ExcelWorkbookProtectionOptions Workbook { get; } = new() { ProtectStructure = false };
    internal ExcelWorkbookWriteReservationOptions Reservation { get; } = new();
    internal bool HasWorkbook { get; private set; }
    internal bool HasReservation { get; private set; }
    internal PdfStandardEncryptionOptions? Pdf { get; private set; }

    internal static OfficeProtectionOptionBinding Capture(EngineIntrinsics engine, OfficeProtectionFamily family,
        object[]? values, string? openPassword, string? permissionsPassword) {
        var result = new OfficeProtectionOptionBinding();
        bool typedPdf = false, typedWorkbook = false, typedReservation = false, typedExcel = false;
        var shellExcel = new OfficeProtectionExcelOptions();
        foreach (object item in values ?? Array.Empty<object>()) {
            object? value = Unwrap(item);
            if (value is IDictionary) continue;
            switch (value) {
                case PdfStandardEncryptionOptions pdf when family == OfficeProtectionFamily.Pdf:
                    if (typedPdf) throw new PSArgumentException("Provide at most one typed PDF encryption options object.");
                    result.Pdf = pdf.Clone();
                    typedPdf = true;
                    break;
                case ExcelWorkbookProtectionOptions workbook when family == OfficeProtectionFamily.Excel:
                    if (typedWorkbook) throw new PSArgumentException("Provide at most one workbook protection options object.");
                    CopyWorkbook(workbook, result.Workbook);
                    result.HasWorkbook = typedWorkbook = true;
                    break;
                case ExcelWorkbookWriteReservationOptions reservation when family == OfficeProtectionFamily.Excel:
                    if (typedReservation) throw new PSArgumentException("Provide at most one write-reservation options object.");
                    CopyReservation(reservation, result.Reservation);
                    result.HasReservation = typedReservation = true;
                    break;
                case OfficeProtectionExcelOptions excel when family == OfficeProtectionFamily.Excel:
                    if (typedExcel) throw new PSArgumentException("Provide at most one Excel command options object.");
                    shellExcel.LegacyPasswordHash = excel.LegacyPasswordHash;
                    shellExcel.ModifyPasswordHash = excel.ModifyPasswordHash;
                    shellExcel.ReservedBy = excel.ReservedBy;
                    typedExcel = true;
                    break;
                default:
                    throw new PSArgumentException("Options must belong to the selected protection format: " + family + ".");
            }
        }
        // Construct the native PDF contract with the real command password, never a dummy password.
        if (family == OfficeProtectionFamily.Pdf && openPassword != null) {
            result.Pdf ??= new PdfStandardEncryptionOptions(openPassword);
            result.Pdf.UserPassword = openPassword;
        }
        if (family == OfficeProtectionFamily.Pdf && permissionsPassword != null) {
            var supplied = result.Pdf;
            result.Pdf = PdfStandardEncryptionOptions.ForPermissions(permissionsPassword,
                supplied?.AllowedPermissions ?? PdfStandardPermissions.All);
            if (supplied != null) {
                result.Pdf.Permissions = supplied.Permissions;
                result.Pdf.Algorithm = supplied.Algorithm;
                result.Pdf.EncryptMetadata = supplied.EncryptMetadata;
                result.Pdf.AesCryptographyProvider = supplied.AesCryptographyProvider;
            }
        }
        foreach (object item in values ?? Array.Empty<object>()) {
            if (Unwrap(item) is not IDictionary table) continue;
            if (family == OfficeProtectionFamily.Pdf) {
                if (result.Pdf == null)
                    throw new PSArgumentException("PDF Options require Encrypt with Password, or a PDF Restriction with an owner Password or RestrictionPassword.");
                var overrides = new OfficeProtectionPdfOptions {
                    Algorithm = result.Pdf.Algorithm,
                    EncryptMetadata = result.Pdf.EncryptMetadata,
                    AesCryptographyProvider = result.Pdf.AesCryptographyProvider
                };
                Bind(engine, overrides, table);
                if (overrides.Permissions.HasValue && overrides.AllowedPermissions.HasValue)
                    throw new PSArgumentException("Specify Permissions or AllowedPermissions, not both in one PDF options table.");
                if (overrides.AllowedPermissions.HasValue &&
                    (overrides.AllowedPermissions.Value & ~PdfStandardPermissions.All) != 0)
                    throw new PSArgumentException("Invalid PDF AllowedPermissions flags.");
                overrides.Apply(result.Pdf);
            } else if (family == OfficeProtectionFamily.Excel) {
                string[] keys = table.Keys.Cast<object>().Select(key => key.ToString() ?? string.Empty).ToArray();
                if (keys.Any(key => key.Equals("ProtectStructure", StringComparison.OrdinalIgnoreCase) || key.Equals("ProtectWindows", StringComparison.OrdinalIgnoreCase))) {
                    Bind(engine, result.Workbook, table);
                    result.HasWorkbook = true;
                } else if (keys.Any(key => key.Equals("ReadOnlyRecommended", StringComparison.OrdinalIgnoreCase) || key.Equals("UserName", StringComparison.OrdinalIgnoreCase))) {
                    Bind(engine, result.Reservation, table);
                    result.HasReservation = true;
                } else Bind(engine, shellExcel, table);
            } else {
                throw new PSArgumentException("No format protection Options are supported for " + family + ".");
            }
        }
        if (shellExcel.LegacyPasswordHash != null) {
            result.Workbook.LegacyPasswordHash = shellExcel.LegacyPasswordHash;
            if (!result.HasWorkbook) result.Workbook.ProtectStructure = true;
            result.HasWorkbook = true;
        }
        if (shellExcel.ModifyPasswordHash != null || shellExcel.ReservedBy != null) {
            if (shellExcel.ModifyPasswordHash != null) result.Reservation.LegacyPasswordHash = shellExcel.ModifyPasswordHash;
            if (shellExcel.ReservedBy != null) result.Reservation.UserName = shellExcel.ReservedBy;
            result.Reservation.ReadOnlyRecommended = true;
            result.HasReservation = true;
        }
        ValidateHash(result.Workbook.LegacyPasswordHash);
        ValidateHash(result.Reservation.LegacyPasswordHash);
        if (result.Pdf != null && !Enum.IsDefined(typeof(PdfStandardEncryptionAlgorithm), result.Pdf.Algorithm))
            throw new PSArgumentException("Invalid PDF encryption Algorithm.");
        return result;
    }

    /// <summary>Combines explicit restrictions with copied native options before any document mutation.</summary>
    internal void PrepareExcel(OfficeProtectionRestriction[] restrictions, string? password) {
        foreach (var restriction in restrictions) {
            if (restriction == OfficeProtectionRestriction.Structure) Workbook.ProtectStructure = true;
            if (restriction == OfficeProtectionRestriction.Windows) Workbook.ProtectWindows = true;
            if (restriction == OfficeProtectionRestriction.ReadOnlyRecommended) Reservation.ReadOnlyRecommended = true;
        }
        HasWorkbook |= Workbook.ProtectStructure || Workbook.ProtectWindows;
        HasReservation |= Reservation.ReadOnlyRecommended;
        if (HasWorkbook) {
            if (password != null) Workbook.Password = password;
            if (!Workbook.ProtectStructure && !Workbook.ProtectWindows)
                throw new PSArgumentException("Workbook protection options must select Structure or Windows.");
        }
        if (HasReservation) {
            if (password != null) Reservation.Password = password;
            if (!Reservation.ReadOnlyRecommended && string.IsNullOrEmpty(Reservation.Password) && string.IsNullOrWhiteSpace(Reservation.LegacyPasswordHash))
                throw new PSArgumentException("Write-reservation options must recommend read-only opening or supply a reservation password/hash.");
        }
    }

    internal static object? Unwrap(object? value) {
        while (value is PSObject wrapper && !ReferenceEquals(value, wrapper.BaseObject)) value = wrapper.BaseObject;
        return value;
    }

    private static void Bind<T>(EngineIntrinsics engine, T target, IDictionary table) where T : class =>
        new OfficeOptionsTransformationAttribute(typeof(T)).ApplyTo(engine, target, table);

    private static void ValidateHash(string? hash) {
        if (string.IsNullOrWhiteSpace(hash)) return;
        string value = hash!.Trim();
        if (value.Length > 4 || value.Any(character => !Uri.IsHexDigit(character)))
            throw new PSArgumentException("Legacy protection hash must be hexadecimal with at most four characters.");
    }

    private static void CopyWorkbook(ExcelWorkbookProtectionOptions source, ExcelWorkbookProtectionOptions target) {
        target.ProtectStructure = source.ProtectStructure;
        target.ProtectWindows = source.ProtectWindows;
        target.Password = source.Password;
        target.LegacyPasswordHash = source.LegacyPasswordHash;
    }

    private static void CopyReservation(ExcelWorkbookWriteReservationOptions source, ExcelWorkbookWriteReservationOptions target) {
        target.ReadOnlyRecommended = source.ReadOnlyRecommended;
        target.UserName = source.UserName;
        target.Password = source.Password;
        target.LegacyPasswordHash = source.LegacyPasswordHash;
    }
}
