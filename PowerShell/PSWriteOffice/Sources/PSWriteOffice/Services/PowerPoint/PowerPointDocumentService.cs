using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Thin lifecycle adapters for PowerShell commands; native paths and encryption remain with OfficeIMO.</summary>
public static partial class PowerPointDocumentService {
    /// <summary>Creates a live presentation with explicit persistence and an associated path.</summary>
    public static PowerPointPresentation CreatePresentation(string filePath) =>
        PowerPointPresentation.Create(Path.GetFullPath(filePath), new PowerPointCreateOptions { PersistenceMode = DocumentPersistenceMode.Explicit });

    /// <summary>Loads a presentation for explicit editing, preserving native source security.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath, string? password = null) =>
        LoadPresentation(filePath, password, readOnly: false);

    /// <summary>Loads a presentation with the requested access mode.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath, bool readOnly) =>
        LoadPresentation(filePath, password: null, readOnly);

    /// <summary>Loads through the native password and access policy without a separate decrypted-document registry.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath, string? password, bool readOnly) =>
        PowerPointPresentation.Load(Path.GetFullPath(filePath), new PowerPointLoadOptions {
            Password = password,
            AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });

    /// <summary>Returns the native document's associated destination.</summary>
    public static string? GetAssociatedPath(PowerPointPresentation presentation) =>
        (presentation ?? throw new ArgumentNullException(nameof(presentation))).FilePath;

    /// <summary>Saves through the native owner without closing the document.</summary>
    public static void SavePresentation(PowerPointPresentation presentation, bool show, string? password = null) =>
        SavePresentation(presentation, show, password, filePath: null);

    /// <summary>Saves to the associated or supplied path, retaining loaded encryption unless a password override is supplied.</summary>
    public static string SavePresentation(PowerPointPresentation presentation, bool show, string? password, string? filePath) {
        EnsureExplicitPersistence(presentation);
        string? destination = filePath ?? presentation.FilePath;
        if (string.IsNullOrWhiteSpace(destination)) throw new InvalidOperationException("Supply a destination path before saving the presentation.");
        destination = Path.GetFullPath(destination!);
        var options = new PowerPointSaveOptions();
        if (!string.IsNullOrEmpty(password)) options.Password = password;
        presentation.Save(destination, options);
        if (show) OfficeShow.Open(destination);
        return destination;
    }

    /// <summary>Closes a presentation, optionally performing an explicitly requested native save first.</summary>
    public static void ClosePresentation(PowerPointPresentation presentation, bool save, bool show, string? password = null) =>
        ClosePresentation(presentation, save, show, password, filePath: null);

    /// <summary>Closes an explicitly persisted presentation after any requested save and show operation.</summary>
    public static void ClosePresentation(PowerPointPresentation presentation, bool save, bool show, string? password, string? filePath) {
        EnsureExplicitPersistence(presentation);
        string? destination = save || show ? SavePresentation(presentation, show: false, password, filePath) : null;
        presentation.Dispose();
        if (show && destination != null) OfficeShow.Open(destination);
    }

    private static void EnsureExplicitPersistence(PowerPointPresentation presentation) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        if (presentation.PersistenceMode != DocumentPersistenceMode.Explicit)
            throw new InvalidOperationException("Use DocumentPersistenceMode.Explicit before saving or closing a presentation through PSWriteOffice.");
    }
}
