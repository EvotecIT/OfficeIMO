using System.Threading;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

public sealed partial class IWorkNativeEditSession {
    /// <summary>Replaces one plain paragraph in a direct Keynote text box without moving native offsets.</summary>
    /// <remarks>The identity is the text box's source identity and the paragraph number is one-based.
    /// Replacement text retains both UTF-16 length and UTF-8 byte length. Only complete modern content and
    /// unique, unlocked freeform shapes with their own storage are admitted. Lists, hyperlinks, inline objects,
    /// shared/inherited storage and nested Index.zip mutations are rejected. Styles and geometry remain intact.
    /// Edited saves require explicit acceptance of retained, unrefreshed native visual caches.</remarks>
    public void SetKeynoteTextBoxParagraphText(IWorkObjectIdentity textBoxIdentity, int paragraph, string text,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (textBoxIdentity == null) throw new ArgumentNullException(nameof(textBoxIdentity));
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (paragraph < 1) throw new ArgumentOutOfRangeException(nameof(paragraph));
        if (_source.Kind != IWorkDocumentKind.Keynote) throw new InvalidOperationException("Text-box editing requires a Keynote source.");
        if (_source.ContainerEntries.Any(entry => entry.Path == "Index.zip"))
            throw new NotSupportedException("Edits to nested Index.zip sources are not supported.");
        ValidateParagraphReplacementText(text);
        var view = _source.WithCancellation(cancellationToken); var projection = view.ReadKeynote();
        if (!projection.HasEditableContent || projection.SourceDeclarationIssues.Count != 0 || projection.SourceReferenceIssues.Count != 0)
            throw new NotSupportedException("Text-box editing requires complete modern Keynote content.");
        var boxes = projection.Slides.SelectMany(slide => slide.Drawables
            .Where(item => item.TextBox != null).Select(item => (Slide: slide, Box: item.TextBox!))).ToArray();
        var matches = boxes.Where(item => SameIdentity(item.Box.SourceIdentity, textBoxIdentity)).Take(2).ToArray();
        if (matches.Length != 1)
            throw new ArgumentException("The identity must resolve to exactly one direct text box in this source.", nameof(textBoxIdentity));
        var selected = matches[0]; var content = selected.Box.Content;
        if (!content.IsComplete || selected.Box.Hyperlink != null || content.SourceIdentity == null
            || content.Paragraphs.Any(item => item.Runs.Any(run => run.InlineObject != null)))
            throw new NotSupportedException("Text-box editing requires complete content without hyperlinks or inline objects.");
        var record = view.Index.Find(textBoxIdentity.RecordIdentifier);
        if (record == null || record.PayloadIndex != 0 || record.MessageType != 2011 || !SameIdentity(new(record), textBoxIdentity))
            throw new NotSupportedException("Only primary freeform text shapes are supported.");
        var drawable = IWorkDrawingReader.DrawableMessage(view.Index, record, out bool complete);
        if (!complete || drawable == null || view.Index.Dereference(drawable, 2)?.Identifier != selected.Slide.SourceIdentity?.RecordIdentifier
            || drawable.FieldCount(5) > 1 || drawable.HasUnexpectedWireKind(5, IWorkWireKind.Varint) || drawable.GetUnsigned(5) is > 0)
            throw new NotSupportedException("Only unlocked direct slide text shapes are supported.");
        var storage = content.SourceIdentity;
        if (boxes.Count(item => SameIdentity(item.Box.Content.SourceIdentity, storage)) != 1
            || projection.Slides.Any(slide => SameIdentity(slide.PresenterNoteContent.SourceIdentity, storage)))
            throw new NotSupportedException("Shared text storage is not supported.");
        var references = new IWorkSourceReferenceIssueCollector(view);
        foreach (var candidate in view.Records.Where(item => item.PayloadIndex == 0 && item.MessageType is 2011 or 7)) {
            cancellationToken.ThrowIfCancellationRequested();
            var other = IWorkKeynoteReader.DrawableStorage(view.Index, candidate, references, out bool storageComplete);
            if (!storageComplete || other != null && other.Identifier == storage.RecordIdentifier && candidate.Identifier != record.Identifier)
                throw new NotSupportedException("Text-shape storage ownership must be unambiguous and unshared.");
        }
        ReplaceNativeParagraphText(view, content, paragraph, text, cancellationToken);
    }
}
