using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>One Keynote slide recovered in presentation order.</summary>
public sealed class IWorkKeynoteSlide {
    internal IWorkKeynoteSlide(int index, string title, IReadOnlyList<string> body,
        string presenterNotes, bool isSkipped) {
        Index = index;
        Title = title;
        Body = body;
        PresenterNotes = presenterNotes;
        IsSkipped = isSkipped;
    }

    /// <summary>Gets the one-based slide position.</summary>
    public int Index { get; }
    /// <summary>Gets title-placeholder text.</summary>
    public string Title { get; }
    /// <summary>Gets remaining editable text blocks.</summary>
    public IReadOnlyList<string> Body { get; }
    /// <summary>Gets presenter-note text.</summary>
    public string PresenterNotes { get; }
    /// <summary>Gets whether the source slide is skipped in the show.</summary>
    public bool IsSkipped { get; }
}

/// <summary>Read-only Keynote structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkKeynoteProjection {
    private readonly IWorkSourceDocument _source;
    private readonly IReadOnlyCollection<ulong> _recognizedIdentifiers;

    internal IWorkKeynoteProjection(IWorkSourceDocument source, IReadOnlyList<IWorkKeynoteSlide> slides,
        IReadOnlyCollection<ulong> recognizedIdentifiers, IReadOnlyList<IWorkDiagnostic> diagnostics) {
        _source = source;
        Slides = slides;
        _recognizedIdentifiers = recognizedIdentifiers;
        Diagnostics = diagnostics;
    }

    /// <summary>Gets presented slides in source order.</summary>
    public IReadOnlyList<IWorkKeynoteSlide> Slides { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether at least one editable slide was recovered.</summary>
    public bool HasEditableContent => Slides.Count > 0;

    /// <summary>Creates an import report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, _recognizedIdentifiers, Diagnostics, preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Slides.Count + Slides.Sum(slide => slide.Body.Count + (slide.Title.Length > 0 ? 1 : 0)
                    + (slide.PresenterNotes.Length > 0 ? 1 : 0)));
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent) {
            throw new InvalidOperationException("Editable Keynote content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Keynote package into a bounded semantic source projection, or returns a diagnostic-only projection in visual-only mode.</summary>
    public IWorkKeynoteProjection ReadKeynote() {
        if (Kind != IWorkDocumentKind.Keynote) throw new InvalidOperationException($"The source is {Kind}, not Keynote.");
        if (RequestedImportMode == IWorkImportMode.VisualOnly) {
            return new IWorkKeynoteProjection(this, Array.Empty<IWorkKeynoteSlide>(), Array.Empty<ulong>(),
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped });
        }
        return IWorkKeynoteReader.Read(this);
    }
}

internal static class IWorkKeynoteReader {
    private const uint DocumentArchive = 1;
    private const uint TextStorageArchive = 2001;

    internal static IWorkKeynoteProjection Read(IWorkSourceDocument source) {
        var recognized = new HashSet<ulong>();
        var diagnostics = new List<IWorkDiagnostic>();
        var slides = new List<IWorkKeynoteSlide>();
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.FirstOfType(DocumentArchive);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_DOCUMENT_MISSING",
                "No supported Keynote document root was found; editable reconstruction is unavailable."));
            return new IWorkKeynoteProjection(source, slides, recognized, diagnostics);
        }
        recognized.Add(document.Identifier);
        IWorkArchiveRecord? show = index.Dereference(index.Message(document), 2);
        if (show == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SHOW_MISSING",
                "The Keynote document root does not reference a supported show object.", document.EntryPath, document.Identifier));
            return new IWorkKeynoteProjection(source, slides, recognized, diagnostics);
        }
        recognized.Add(show.Identifier);
        IWorkWireMessage? slideTree = IWorkObjectIndex.TryGetMessage(index.Message(show), 3);
        if (slideTree == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_KEYNOTE_SLIDE_TREE_MISSING",
                "The Keynote show does not contain a supported slide tree.", show.EntryPath, show.Identifier));
            return new IWorkKeynoteProjection(source, slides, recognized, diagnostics);
        }

        int position = 0;
        foreach (IWorkArchiveRecord node in index.DereferenceAll(slideTree, 2)) {
            position++;
            recognized.Add(node.Identifier);
            IWorkWireMessage nodeMessage = index.Message(node);
            bool skipped = nodeMessage.GetUnsigned(4) == 1;
            IWorkArchiveRecord? slide = index.Dereference(nodeMessage, 2);
            if (slide == null) {
                slides.Add(new IWorkKeynoteSlide(position, string.Empty, Array.Empty<string>(), string.Empty, skipped));
                continue;
            }
            recognized.Add(slide.Identifier);
            slides.Add(ReadSlide(index, slide, position, skipped, recognized));
        }
        return new IWorkKeynoteProjection(source, slides, recognized, diagnostics);
    }

    private static IWorkKeynoteSlide ReadSlide(IWorkObjectIndex index, IWorkArchiveRecord slide,
        int position, bool skipped, HashSet<ulong> recognized) {
        IWorkWireMessage message = index.Message(slide);
        IWorkArchiveRecord? titlePlaceholder = index.Dereference(message, 5);
        var candidates = new List<IWorkArchiveRecord>();
        foreach (int field in new[] { 5, 6, 7, 42 }) {
            foreach (IWorkArchiveRecord candidate in index.DereferenceAll(message, field)) {
                if (!candidates.Any(existing => existing.Identifier == candidate.Identifier)) candidates.Add(candidate);
            }
        }

        string title = string.Empty;
        var body = new List<string>();
        var seenStorages = new HashSet<ulong>();
        foreach (IWorkArchiveRecord drawable in candidates) {
            IWorkArchiveRecord? storage = DrawableStorage(index, drawable);
            if (storage == null || storage.MessageType != TextStorageArchive || !seenStorages.Add(storage.Identifier)) continue;
            string text = IWorkPagesReader.StorageText(index.Message(storage)).Trim();
            if (text.Length == 0) continue;
            recognized.Add(drawable.Identifier);
            recognized.Add(storage.Identifier);
            if (titlePlaceholder != null && drawable.Identifier == titlePlaceholder.Identifier && title.Length == 0) title = text;
            else body.Add(text);
        }

        string notes = string.Empty;
        IWorkArchiveRecord? note = index.Dereference(message, 27);
        if (note != null) {
            IWorkArchiveRecord? storage = index.Dereference(index.Message(note), 1);
            if (storage != null && storage.MessageType == TextStorageArchive) {
                notes = IWorkPagesReader.StorageText(index.Message(storage)).TrimStart('\n');
                recognized.Add(note.Identifier);
                recognized.Add(storage.Identifier);
            }
        }
        return new IWorkKeynoteSlide(position, title, body, notes, skipped);
    }

    private static IWorkArchiveRecord? DrawableStorage(IWorkObjectIndex index, IWorkArchiveRecord drawable) {
        IWorkWireMessage message = index.Message(drawable);
        IWorkArchiveRecord? direct = index.Dereference(message, 2);
        if (direct != null && direct.MessageType == TextStorageArchive) return direct;
        IWorkWireMessage? super = IWorkObjectIndex.TryGetMessage(message, 1);
        if (super == null) return null;
        IWorkArchiveRecord? nested = index.Dereference(super, 2);
        return nested != null && nested.MessageType == TextStorageArchive ? nested : null;
    }
}
