using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

/// <summary>An editable publication and its retained manuscript review state. Instances are mutable and not thread-safe.</summary>
public sealed partial class BookProject {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _diagnostics;
    private EpubPublication _publication;
    private byte[]? _undo, _redo;
    private BookProject(EpubPublication publication, IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics, bool acknowledged) {
        _publication = publication;
        _diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        ImportLossAcknowledged = acknowledged;
    }
    /// <summary>Canonical EPUB authoring model backing this project.</summary>
    public EpubPublication Publication => _publication;
    /// <summary>Whether one validated package edit can be undone. History is session-only and bounded to 128 MiB per snapshot.</summary>
    public bool CanUndo => _undo != null;
    /// <summary>Whether the last package undo can be redone.</summary>
    public bool CanRedo => _redo != null;
    /// <summary>Restores the previous validated publication while retaining import review state.</summary>
    public void Undo(CancellationToken cancellationToken = default) {
        byte[] previous = _undo ?? throw new InvalidOperationException("No package edit is available to undo.");
        byte[] current = _publication.Write(cancellationToken: cancellationToken).Bytes;
        using var stream = new MemoryStream(previous, false);
        var restored = EpubPublication.Load(stream, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _publication = restored; _undo = null; _redo = current.LongLength <= 128L * 1024 * 1024 ? current : null;
    }
    /// <summary>Restores the publication that was replaced by the last undo.</summary>
    public void Redo(CancellationToken cancellationToken = default) {
        byte[] next = _redo ?? throw new InvalidOperationException("No package edit is available to redo.");
        byte[] current = _publication.Write(cancellationToken: cancellationToken).Bytes;
        using var stream = new MemoryStream(next, false);
        var restored = EpubPublication.Load(stream, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        _publication = restored; _redo = null; _undo = current.LongLength <= 128L * 1024 * 1024 ? current : null;
    }
    /// <summary>Retained category-preserving manuscript import diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> ImportDiagnostics => _diagnostics;
    /// <summary>Whether the author has reviewed and accepted non-fatal import losses.</summary>
    public bool ImportLossAcknowledged { get; private set; }
    /// <summary>Whether import review permits export. Native writer validation still runs for every export.</summary>
    public bool CanExport => !_diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure) &&
        (ImportLossAcknowledged || !_diagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None));
    /// <summary>Creates a reflowable EPUB project with an editable first chapter.</summary>
    public static BookProject Create(string title, string language = "en") {
        var publication = EpubPublication.Create(title, language);
        publication.AddChapter("chapter-1", "EPUB/text/chapter-1.xhtml", "Chapter 1", "<h1>Chapter 1</h1><p>Write your book here.</p>");
        return new BookProject(publication, [], false);
    }
    /// <summary>Updates primary package metadata atomically, retaining other values and refinements. A blank creator retains the current value.</summary>
    public void SetMetadata(string title, string language, string creator, CancellationToken cancellationToken = default) =>
        Mutate(publication => { publication.Title = title; publication.Language = language; if (!string.IsNullOrWhiteSpace(creator)) publication.Creator = creator; }, cancellationToken);
    /// <summary>Replaces a chapter's XHTML body while retaining its document head and package declarations.</summary>
    public void SetChapterBody(int index, string bodyXml, CancellationToken cancellationToken = default) =>
        Mutate(publication => {
            string id = publication.Spine[index].ManifestId;
            XDocument content = publication.GetContentXml(id);
            XElement body = ParseEditorBody(bodyXml);
            content.Root!.Element(body.Name)!.ReplaceWith(body);
            publication.SetContentXml(id, content);
        }, cancellationToken);
    /// <summary>Starts a project from the shared manuscript import result.</summary>
    public static BookProject FromImport(EpubManuscriptResult result) {
        ArgumentNullException.ThrowIfNull(result);
        return new BookProject(result.Publication, result.Report.FidelityDiagnostics, false);
    }
    /// <summary>Opens an existing EPUB for package-preserving editing without claiming that its content was converted.</summary>
    public static BookProject FromEpub(byte[] bytes, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        using var source = new MemoryStream(bytes, false);
        return new BookProject(EpubPublication.Load(source, cancellationToken: cancellationToken), [], false);
    }
    /// <summary>Records author acceptance of import approximations and omissions. Failed imports require source repair and re-import.</summary>
    public void AcknowledgeImportLoss() {
        if (_diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure))
            throw new InvalidOperationException("Failed manuscript dependencies must be repaired before export.");
        ImportLossAcknowledged = true;
    }
    /// <summary>Exports a complete EPUB after import review and canonical writer validation.</summary>
    public EpubWriteResult Export(CancellationToken cancellationToken = default) {
        if (!CanExport) throw new InvalidOperationException("Review the import report before exporting this book.");
        return _publication.Write(cancellationToken: cancellationToken);
    }
    /// <summary>Moves a chapter and its top-level navigation entry atomically, retaining nested navigation.</summary>
    public void MoveChapter(int fromIndex, int toIndex, CancellationToken cancellationToken = default) {
        Mutate(proposed => {
            EpubDocument reading = proposed.Read(cancellationToken: cancellationToken);
            string path = proposed.Manifest.Single(item => item.Id == proposed.Spine[fromIndex].ManifestId).Reference.ContainerPath!;
            string targetPath = proposed.Manifest.Single(item => item.Id == proposed.Spine[toIndex].ManifestId).Reference.ContainerPath!;
            var roots = reading.TableOfContents.ToList();
            int rootIndex = roots.FindIndex(item => item.Target == path);
            int targetIndex = roots.FindIndex(item => item.Target == targetPath);
            proposed.MoveSpineItem(fromIndex, toIndex);
            if (rootIndex >= 0 && targetIndex >= 0 && rootIndex != targetIndex) {
                EpubNavigationItem moved = roots[rootIndex]; roots.RemoveAt(rootIndex); roots.Insert(targetIndex, moved);
                proposed.SetNavigation(roots.Select(item => ToNavigation(item)));
            }
        }, cancellationToken);
    }
    /// <summary>Renames a chapter's document title, first matching heading and matching primary navigation label atomically.</summary>
    public void RenameChapter(int index, string title, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        Mutate(proposed => RenameChapterCore(proposed, index, title, cancellationToken), cancellationToken);
    }
    private static void RenameChapterCore(EpubPublication proposed, int index, string title, CancellationToken cancellationToken) {
            ArgumentException.ThrowIfNullOrWhiteSpace(title);
            string id = proposed.Spine[index].ManifestId;
            string path = proposed.Manifest.Single(item => item.Id == id).Reference.ContainerPath!;
            EpubDocument reading = proposed.Read(cancellationToken: cancellationToken);
            XDocument content = proposed.GetContentXml(id);
            XNamespace html = "http://www.w3.org/1999/xhtml";
            XElement documentTitle = content.Root?.Element(html + "head")?.Element(html + "title")
                ?? throw new NotSupportedException("Chapter title editing requires XHTML content.");
            string previousTitle = documentTitle.Value;
            XElement? heading = content.Descendants().FirstOrDefault(element => element.Name.Namespace == html &&
                new[] { "h1", "h2", "h3", "h4", "h5", "h6" }.Contains(element.Name.LocalName));
            if (heading?.Value.Trim() == previousTitle && !heading.Elements().Any()) heading.Value = title;
            documentTitle.Value = title;
            proposed.SetContentXml(id, content);
            proposed.SetNavigation(reading.TableOfContents.Select(item => ToNavigation(item, path, title)));
    }
    /// <summary>Changes the project typography stylesheet without replacing imported source CSS.</summary>
    public void SetStylesheet(string css, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(css);
        Mutate(proposed => ApplyStylesheet(proposed, css), cancellationToken);
    }
    private static void ApplyStylesheet(EpubPublication proposed, string css) {
            const string id = "book-project-style";
            var existing = proposed.Manifest.SingleOrDefault(item => item.Id == id);
            string path = existing?.Reference.ContainerPath ?? "EPUB/styles/project.css";
            if (existing == null) proposed.AddStylesheet(id, path, css);
            else {
                if (!string.Equals(existing.MediaType, "text/css", StringComparison.OrdinalIgnoreCase) || existing.Reference.ContainerPath == null)
                    throw new InvalidDataException("The project stylesheet identifier is already assigned to a different resource kind.");
                proposed.UpdateResource(id, new System.Text.UTF8Encoding(false, true).GetBytes(css));
            }
            foreach (EpubSpineItem position in proposed.Spine) {
                if (proposed.Manifest.Single(item => item.Id == position.ManifestId).MediaType != "application/xhtml+xml") continue;
                XDocument content = proposed.GetContentXml(position.ManifestId);
                XNamespace html = "http://www.w3.org/1999/xhtml";
                XElement head = content.Root!.Element(html + "head")!;
                string owner = proposed.Manifest.Single(item => item.Id == position.ManifestId).Reference.ContainerPath!;
                string? baseHref = head.Elements(html + "base").Select(element => (string?)element.Attribute("href")).FirstOrDefault();
                EpubReference effectiveBase = EpubReference.Resolve(owner, baseHref, "__book_relative_base__");
                if (effectiveBase.Kind != EpubReferenceKind.Container)
                    throw new NotSupportedException("Project stylesheet editing requires a local chapter base URL.");
                string href = RelativePackageHref(effectiveBase.ContainerPath!, path);
                if (!head.Elements(html + "link").Any(link => (string?)link.Attribute("href") == href))
                    head.Add(new XElement(html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", href), new XAttribute("type", "text/css")));
                proposed.SetContentXml(position.ManifestId, content);
            }
    }
    /// <summary>Applies a complete editor draft in one transaction; failure retains the previous publication.</summary>
    public void ApplyEdits(BookProjectEdits edits, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(edits);
        var bodies = edits.ChapterBodies?.ToArray() ?? [];
        var titles = edits.ChapterTitles?.ToArray() ?? [];
        string? title = edits.Title, language = edits.Language, creator = edits.Creator, css = edits.Stylesheet;
        Mutate(proposed => {
            if (title != null && title != proposed.Title) proposed.Title = title;
            if (language != null && language != proposed.Language) proposed.Language = language;
            if (!string.IsNullOrWhiteSpace(creator) && creator != proposed.Creator) proposed.Creator = creator;
            foreach (var replacement in bodies) {
                XDocument content = proposed.GetContentXml(replacement.Key);
                XElement body = ParseEditorBody(replacement.Value);
                content.Root!.Element(body.Name)!.ReplaceWith(body);
                proposed.SetContentXml(replacement.Key, content);
            }
            foreach (var replacement in titles) {
                int index = Enumerable.Range(0, proposed.Spine.Count).Single(position => proposed.Spine[position].ManifestId == replacement.Key);
                RenameChapterCore(proposed, index, replacement.Value, cancellationToken);
            }
            if (css != null) ApplyStylesheet(proposed, css);
        }, cancellationToken);
    }
    /// <summary>Selects newly supplied image bytes as the package cover, retaining earlier resources and content references.</summary>
    public void SetCoverImage(byte[] bytes, string mediaType, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        ArgumentException.ThrowIfNullOrWhiteSpace(mediaType);
        Mutate(proposed => {
            string extension = mediaType.ToLowerInvariant() switch {
                "image/png" => ".png", "image/jpeg" => ".jpg", "image/gif" => ".gif", "image/svg+xml" => ".svg",
                _ => throw new NotSupportedException("Cover import supports PNG, JPEG, GIF and static SVG.")
            };
            int next = 1;
            string id, path;
            do {
                id = "book-cover-" + next;
                path = "EPUB/resources/book-cover-" + next++ + extension;
            } while (proposed.Manifest.Any(item => item.Id == id) || proposed.EntryPaths.Contains(path));
            proposed.AddResource(id, path, mediaType.ToLowerInvariant(), bytes);
            proposed.SetCoverImage(id);
        }, cancellationToken);
    }
    /// <summary>Renders one current chapter through the existing EPUB image owner, using retained package resources only.</summary>
    public IReadOnlyList<OfficeImageExportResult> PreviewChapter(int index, double width = 816,
        CancellationToken cancellationToken = default) {
        EpubDocument source = _publication.Read(new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true, MaxChapters = _publication.Spine.Count }, cancellationToken);
        return source.ExportImages(OfficeImageExportFormat.Png, new EpubImageExportOptions {
            ChapterIndex = index, ChapterCount = 1, ViewportWidth = width, Mode = HtmlRenderMode.Continuous,
            MaximumOutputCount = 1
        }, cancellationToken);
    }
    private void Mutate(Action<EpubPublication> mutation, CancellationToken token) {
        byte[] bytes = _publication.Write(cancellationToken: token).Bytes;
        using var stream = new MemoryStream(bytes, false);
        EpubPublication proposed = EpubPublication.Load(stream, cancellationToken: token);
        mutation(proposed);
        byte[] validated = proposed.Write(cancellationToken: token).Bytes;
        token.ThrowIfCancellationRequested();
        if (!bytes.SequenceEqual(validated)) { _undo = bytes.LongLength <= 128L * 1024 * 1024 ? bytes : null; _redo = null; }
        _publication = proposed;
    }
    private static EpubNavigationEntry ToNavigation(EpubNavigationItem item, string? renamePath = null, string? title = null) =>
        new EpubNavigationEntry(item.Target == renamePath ? title! : item.Label,
            (item.Target ?? throw new InvalidDataException("Navigation entry has no resolved target.")) +
                (item.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(item.Fragment)),
            item.Children.Select(child => ToNavigation(child)), item.SemanticType);
}
