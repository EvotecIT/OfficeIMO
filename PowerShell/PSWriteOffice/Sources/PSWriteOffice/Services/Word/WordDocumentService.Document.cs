using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

/// <summary>Bridges DSL cmdlets with OfficeIMO.Word document operations.</summary>
public static partial class WordDocumentService
{
    private static readonly AsyncLocal<WeakReference<WordDocument>[]?> TrackedDocuments = new();
    private static readonly ConditionalWeakTable<WordDocument, DocumentAssociation> Associations = new();

    private sealed class DocumentAssociation {
        internal string Path { get; set; } = string.Empty;
        internal bool Encrypted { get; set; }
    }

    /// <summary>Creates an editable detached copy without modifying the template.</summary>
    public static WordDocument CreateFromTemplate(string templatePath) =>
        RegisterDocument(WordDocument.Create(new WordCreateOptions { Template = templatePath }));

    /// <summary>Loads an existing Word document.</summary>
    public static WordDocument LoadDocument(string filePath, bool readOnly, bool autoSave, string? password = null)
    {
        var resolvedPath = Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
        {
            throw new FileNotFoundException($"File {resolvedPath} doesn't exist.", resolvedPath);
        }

        if (!string.IsNullOrEmpty(password))
        {
            var document = RegisterDocument(OfficeEncryptedPackageService.LoadWord(resolvedPath, password!, readOnly, autoSave));
            UpdateSaveAssociation(document, resolvedPath, encrypted: true);
            return document;
        }

        return RegisterDocument(WordDocument.Load(resolvedPath, CreateLoadOptions(readOnly, autoSave)));
    }

    /// <summary>Creates a new Word document at the specified path.</summary>
    public static WordDocument CreateDocument(string filePath, bool autoSave)
    {
        return RegisterDocument(WordDocument.Create(Path.GetFullPath(filePath), new WordCreateOptions
        {
            PersistenceMode = autoSave ? DocumentPersistenceMode.SaveOnDispose : DocumentPersistenceMode.Explicit
        }));
    }

    /// <summary>Creates a new in-memory Word document without creating a package on disk.</summary>
    public static WordDocument CreateInMemoryDocument()
    {
        return RegisterDocument(WordDocument.Create());
    }

    /// <summary>Disposes the Word document.</summary>
    public static void CloseDocument(WordDocument document)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        try
        {
            document.Dispose();
        }
        finally
        {
            Associations.Remove(document);
            UnregisterDocument(document);
        }
    }

    /// <summary>Saves the document, optionally to a new path, and closes it.</summary>
    public static void SaveDocument(WordDocument document, bool show, string? filePath, string? password = null)
    {
        var associatedPath = GetAssociatedPath(document);
        if (string.IsNullOrWhiteSpace(associatedPath) && string.IsNullOrWhiteSpace(filePath))
        {
            throw new InvalidOperationException("No file path provided.");
        }

        if (!string.IsNullOrWhiteSpace(filePath))
        {
            var targetPath = Path.GetFullPath(filePath!);
            if (!string.Equals(targetPath, associatedPath, StringComparison.OrdinalIgnoreCase))
            {
                SaveDocumentToPath(document, targetPath, false, password);
                CloseDocument(document);
                if (show)
                {
                    FileOpenService.Open(targetPath);
                }
                return;
            }
        }

        if (!string.IsNullOrEmpty(password))
        {
            var targetPath = associatedPath!;
            OfficeEncryptedPackageService.SaveWord(document, targetPath, password!, false);
        }
        else
        {
            if (IsEncryptedSource(document))
            {
                throw new InvalidOperationException("Provide -Password when saving a document loaded from an encrypted package.");
            }

            document.Save(associatedPath!);
        }

        var savedPath = associatedPath ?? throw new InvalidOperationException("No saved file path was available.");
        CloseDocument(document);

        if (show)
        {
            FileOpenService.Open(savedPath);
        }
    }

    private static void SaveDocumentToPath(WordDocument document, string path, bool openWord, string? password)
    {
        if (!string.IsNullOrEmpty(password))
        {
            OfficeEncryptedPackageService.SaveWord(document, path, password!, openWord);
            return;
        }

        document.Save(path);
    }

    private static WordLoadOptions CreateLoadOptions(bool readOnly, bool autoSave) => new()
    {
        AccessMode = readOnly ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite,
        PersistenceMode = autoSave ? DocumentPersistenceMode.SaveOnDispose : DocumentPersistenceMode.Explicit
    };

    /// <summary>Returns the most recently tracked Word document for the current runspace.</summary>
    public static WordDocument? GetCurrentTrackedDocument()
    {
        var tracked = GetAliveTrackedDocuments();
        return tracked.Length == 0
            ? null
            : tracked[tracked.Length - 1];
    }

    /// <summary>Returns tracked Word documents for the current runspace.</summary>
    public static IReadOnlyList<WordDocument> GetTrackedDocuments()
    {
        return GetAliveTrackedDocuments();
    }

    private static T RegisterDocument<T>(T document) where T : WordDocument
    {
        var tracked = GetAliveTrackedDocuments();
        TrackedDocuments.Value = tracked
            .Where(existing => !ReferenceEquals(existing, document))
            .Append(document)
            .Select(item => new WeakReference<WordDocument>(item))
            .ToArray();
        return document;
    }

    private static void UnregisterDocument(WordDocument document) {
        TrackedDocuments.Value = GetAliveTrackedDocuments()
            .Where(existing => !ReferenceEquals(existing, document))
            .Select(item => new WeakReference<WordDocument>(item))
            .ToArray();
    }

    private static WordDocument[] GetAliveTrackedDocuments() {
        var alive = (TrackedDocuments.Value ?? Array.Empty<WeakReference<WordDocument>>())
            .Select(reference => reference.TryGetTarget(out var document) ? document : null)
            .Where(document => document != null && !document.IsDisposed)
            .Cast<WordDocument>()
            .ToArray();
        TrackedDocuments.Value = alive.Select(item => new WeakReference<WordDocument>(item)).ToArray();
        return alive;
    }

    internal static string? GetAssociatedPath(WordDocument document) =>
        Associations.TryGetValue(document, out var association) ? association.Path : document.FilePath;

    internal static bool IsEncryptedSource(WordDocument document) =>
        Associations.TryGetValue(document, out var association) && association.Encrypted;

    internal static void UpdateSaveAssociation(WordDocument document, string path, bool encrypted) {
        var association = Associations.GetValue(document, _ => new DocumentAssociation());
        association.Path = Path.GetFullPath(path);
        association.Encrypted = encrypted;
    }
}
