using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherModelState {
    internal PublisherModelState AddPage(bool isMaster, string name, uint? masterId, CancellationToken token, out uint id) {
        token.ThrowIfCancellationRequested();
        if (_publication.Pages.Count >= _decoded.Options.MaximumPages) throw new InvalidDataException("Publisher page limit exceeded.");
        PublisherStoryEditor.ValidatePageName(name, _decoded.Options);
        if (_publication.Pages.Sum(page => (long)page.Name.Length) + name.Length > _decoded.Options.Limits.MaxTextCharacters)
            throw new ArgumentException("Publication page names exceed the configured character limit.", nameof(name));
        return WithPublication(_publication.AddPage(isMaster, name, masterId, out id));
    }

    internal PublisherModelState SetPageMaster(uint pageId, uint? masterId) =>
        WithPublication(_publication.SetPageMaster(pageId, masterId));

    internal PublisherModelState AddTextFrame(uint pageId, double x, double y, double width, double height,
        IReadOnlyList<OfficeRichTextParagraph> paragraphs, CancellationToken token, out uint id) {
        token.ThrowIfCancellationRequested();
        if ((long)SourceObjectCount + _publication.AuthoredObjectCount >= _decoded.Options.Limits.MaxItems)
            throw new InvalidDataException("Publisher current object limit exceeded.");
        return WithPublication(_publication.AddTextFrame(pageId, x, y, width, height, paragraphs, _decoded.Options, token, out id));
    }
}
