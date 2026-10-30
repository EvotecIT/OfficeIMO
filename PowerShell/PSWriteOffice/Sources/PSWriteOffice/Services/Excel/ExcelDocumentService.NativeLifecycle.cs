using System;
using System.Net;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static partial class ExcelDocumentService {
    /// <summary>Rejects disposal or saves that could implicitly publish a live workbook.</summary>
    internal static void EnsureExplicitPersistence(ExcelDocument document) {
        if (document.PersistenceMode != DocumentPersistenceMode.Explicit)
            throw new InvalidOperationException("Use DocumentPersistenceMode.Explicit before saving or closing an Excel workbook through PSWriteOffice.");
    }

    /// <summary>Clones native policy, preserving an omitted password and applying only supplied overrides.</summary>
    internal static ExcelSaveOptions NativeSaveOptions(ExcelSaveOptions? source, bool? allowLoss,
        SecureString? password, bool passwordBound) {
        var options = source?.Clone() ?? new ExcelSaveOptions();
        if (allowLoss.HasValue) options.AllowLoss = allowLoss.Value;
        if (passwordBound) options.Password = password == null ? null : new NetworkCredential(string.Empty, password).Password;
        return options;
    }
}
