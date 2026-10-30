using System;
using System.Net;
using System.Security;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService {
    /// <summary>Rejects implicit file writes when a lifecycle command saves or disposes a live document.</summary>
    internal static void EnsureExplicitPersistence(WordDocument document) {
        if (document.PersistenceMode != OfficeIMO.DocumentPersistenceMode.Explicit)
            throw new InvalidOperationException("Use DocumentPersistenceMode.Explicit before saving or closing a Word document through PSWriteOffice.");
    }

    /// <summary>Creates native save policy while retaining loaded encryption when Password was omitted.</summary>
    internal static WordSaveOptions NativeSaveOptions(bool allowLoss, SecureString? password, bool passwordBound) {
        var options = new WordSaveOptions { AllowLoss = allowLoss };
        if (passwordBound) options.Password = password == null ? null : new NetworkCredential(string.Empty, password).Password;
        return options;
    }
}
