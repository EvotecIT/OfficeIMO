using System;
using System.Collections.Concurrent;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Helper methods bridging DSL cmdlets with OfficeIMO PowerPoint presentations.</summary>
public static class PowerPointDocumentService
{
    private sealed class PresentationAssociation
    {
        internal PresentationAssociation(string path, bool encrypted)
        {
            Path = path;
            Encrypted = encrypted;
        }

        internal string Path { get; }
        internal bool Encrypted { get; }
    }

    private static readonly ConditionalWeakTable<PowerPointPresentation, PresentationAssociation> Presentations = new();

    /// <summary>Creates a new presentation at the specified path.</summary>
    public static PowerPointPresentation CreatePresentation(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        }

        var resolvedPath = Path.GetFullPath(filePath);
        var presentation = PowerPointPresentation.Create(resolvedPath, new PowerPointCreateOptions
        {
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
        Track(presentation, resolvedPath, encrypted: false);
        return presentation;
    }

    /// <summary>Loads an existing presentation.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath, string? password = null) =>
        LoadPresentation(filePath, password, readOnly: false);

    /// <summary>Loads an existing presentation with an explicit access mode.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath, bool readOnly) =>
        LoadPresentation(filePath, password: null, readOnly);

    /// <summary>Loads an existing presentation with password and access-mode options.</summary>
    public static PowerPointPresentation LoadPresentation(string filePath, string? password, bool readOnly)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }

        var presentation = !string.IsNullOrEmpty(password)
            ? OfficeEncryptedPackageService.OpenPowerPoint(resolvedPath, password!, readOnly)
            : PowerPointPresentation.Load(resolvedPath, new PowerPointLoadOptions
            {
                AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
                PersistenceMode = DocumentPersistenceMode.Explicit
            });
        Track(presentation, resolvedPath, encrypted: !string.IsNullOrEmpty(password));
        return presentation;
    }

    /// <summary>Returns the associated path for a presentation, including externally created OfficeIMO instances.</summary>
    public static string? GetAssociatedPath(PowerPointPresentation presentation)
    {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        return Presentations.TryGetValue(presentation, out var association)
            ? association.Path
            : presentation.FilePath;
    }

    /// <summary>Saves without closing and optionally opens the persisted presentation.</summary>
    public static void SavePresentation(PowerPointPresentation presentation, bool show, string? password = null) =>
        SavePresentation(presentation, show, password, filePath: null);

    /// <summary>Saves without closing, optionally to a new path, and returns the associated destination.</summary>
    public static string SavePresentation(PowerPointPresentation presentation, bool show, string? password = null, string? filePath = null)
    {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        EnsureExternalPresentationUsesExplicitPersistence(presentation);
        var associatedPath = GetAssociatedPath(presentation);
        if (string.IsNullOrWhiteSpace(filePath) && string.IsNullOrWhiteSpace(associatedPath))
        {
            throw new InvalidOperationException("No file path provided. Use -Path or open the presentation from disk.");
        }

        var resolvedPath = Path.GetFullPath(string.IsNullOrWhiteSpace(filePath) ? associatedPath! : filePath!);
        string? targetDirectory = Path.GetDirectoryName(resolvedPath);
        if (!string.IsNullOrWhiteSpace(targetDirectory))
        {
            Directory.CreateDirectory(targetDirectory);
        }
        if (!string.IsNullOrEmpty(password))
        {
            OfficeEncryptedPackageService.SavePowerPoint(presentation, resolvedPath, password!);
        }
        else
        {
            bool trackedEncryptedSource = Presentations.TryGetValue(presentation, out var association) &&
                association.Encrypted &&
                PathsIdentifySameFile(resolvedPath, Path.GetFullPath(associatedPath!));
            if (trackedEncryptedSource || IsEncryptedTarget(resolvedPath))
            {
                throw new InvalidOperationException("Provide -Password when saving a presentation loaded from an encrypted package.");
            }

            presentation.Save(resolvedPath);
        }

        Track(presentation, resolvedPath, encrypted: !string.IsNullOrEmpty(password));

        if (show)
        {
            FileOpenService.Open(resolvedPath);
        }

        return resolvedPath;
    }

    /// <summary>Closes a presentation, optionally saving and opening it first.</summary>
    public static void ClosePresentation(PowerPointPresentation presentation, bool save, bool show, string? password = null) =>
        ClosePresentation(presentation, save, show, password, filePath: null);

    /// <summary>Closes a presentation, optionally saving to a new path and opening it first.</summary>
    public static void ClosePresentation(PowerPointPresentation presentation, bool save, bool show, string? password, string? filePath)
    {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        EnsureExternalPresentationUsesExplicitPersistence(presentation);
        string? savedPath = null;
        if (save || show)
        {
            savedPath = SavePresentation(presentation, show: false, password, filePath);
        }

        try
        {
            presentation.Dispose();
        }
        finally
        {
            Presentations.Remove(presentation);
        }

        if (show && savedPath != null)
        {
            FileOpenService.Open(savedPath);
        }
    }

    private static bool IsEncryptedTarget(string path)
    {
        string extension = Path.GetExtension(path);
        if (!extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".pptm", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".potx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".potm", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".ppsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".ppsm", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".ppam", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return OfficeEncryptedPackageService.HasCompoundFileSignature(path);
    }

    private static void EnsureExternalPresentationUsesExplicitPersistence(
        PowerPointPresentation presentation)
    {
        if (!Presentations.TryGetValue(presentation, out _) &&
            presentation.PersistenceMode == DocumentPersistenceMode.SaveOnDispose)
        {
            throw new InvalidOperationException(
                "External presentations using SaveOnDispose cannot be managed safely. " +
                "Create or load the presentation with DocumentPersistenceMode.Explicit before using PSWriteOffice save or close commands.");
        }
    }

    private static bool PathsIdentifySameFile(string left, string right)
    {
        StringComparison comparison = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
                                      RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), comparison);
    }

    private static void Track(PowerPointPresentation presentation, string path, bool encrypted)
    {
        Presentations.Remove(presentation);
        Presentations.Add(presentation, new PresentationAssociation(path, encrypted));
    }
}
