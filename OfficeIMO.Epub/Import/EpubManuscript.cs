using AngleSharp.Dom;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Epub;

/// <summary>Imports inert HTML manuscripts through the shared HTML owner into a reflowable EPUB 3 publication.</summary>
public static partial class EpubManuscript {
    private static readonly XNamespace Xhtml = "http://www.w3.org/1999/xhtml";
    private const string DefaultCss = "body{line-height:1.5}img,svg{max-width:100%;height:auto}table{border-collapse:collapse;max-width:100%}th,td{padding:.25em}pre{white-space:pre-wrap;overflow-wrap:anywhere}a{overflow-wrap:anywhere}";

    /// <summary>Imports a prepared HTML manuscript with embedded resources. Use the asynchronous route for an application resolver.</summary>
    public static EpubManuscriptResult ImportHtml(HtmlConversionDocument manuscript, EpubManuscriptOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (manuscript == null) throw new ArgumentNullException(nameof(manuscript));
        options ??= new EpubManuscriptOptions();
        if (options.ResourceResolver != null) throw new ArgumentException("Use ImportHtmlAsync when supplying a resource resolver.", nameof(options));
        return ImportHtmlAsync(manuscript, options, cancellationToken).GetAwaiter().GetResult();
    }

    /// <summary>Imports HTML and collects its resource dependencies through an explicitly supplied application resolver.</summary>
    public static async Task<EpubManuscriptResult> ImportHtmlAsync(HtmlConversionDocument manuscript, EpubManuscriptOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (manuscript == null) throw new ArgumentNullException(nameof(manuscript));
        options ??= new EpubManuscriptOptions();
        options = options.Clone();
        cancellationToken.ThrowIfCancellationRequested();
        if (options.ChapterHeadingLevel < 0 || options.ChapterHeadingLevel > 6) throw new ArgumentOutOfRangeException(nameof(options.ChapterHeadingLevel));
        var source = manuscript.CreateSourceDocumentForConversion();
        string title = options.Title ?? source.Title ?? string.Empty;
        string language = options.Language ?? source.DocumentElement?.GetAttribute("lang") ?? source.DocumentElement?.GetAttribute("xml:lang") ?? "en";
        var publication = EpubPublication.Create(title, language, options.Identifier, retentionLimits: options.RetentionLimits);
        string? creator = options.Creator ?? source.QuerySelector("meta[name='author']")?.GetAttribute("content");
        if (!string.IsNullOrWhiteSpace(creator)) publication.Creator = creator;
        foreach (string name in new[] { "description", "rights" }) {
            string? value = source.QuerySelector("meta[name='" + name + "']")?.GetAttribute("content");
            if (!string.IsNullOrWhiteSpace(value)) publication.AddDublinCoreMetadata(name, value!);
        }
        var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
        XElement body = ConvertElement(source.Body!, diagnostics, manuscript, cancellationToken)!;
        var chapters = SplitChapters(body, options.ChapterHeadingLevel, title, diagnostics, cancellationToken);
        AssignAnchors(chapters, diagnostics);
        RewriteChapterLinks(chapters, manuscript, diagnostics);
        foreach (Chapter chapter in chapters) {
            try {
                var ids = EpubContentIdentifiers.Collect(chapter.Body, chapter.Path, true, cancellationToken);
                EpubContentIdentifiers.ValidateReferences(chapter.Body, ids, chapter.Path, cancellationToken);
            } catch (InvalidDataException error) {
                AddDiagnostic(diagnostics, "EPUB_IMPORT_ID_REFERENCE_INVALID", error.Message, chapter.Path, OfficeConversionLossKind.Failure);
            }
        }
        foreach (var active in source.QuerySelectorAll("script,iframe,object,embed,form,input,button,select,textarea")) active.Remove();
        List<ImportedStylesheet> sourceStyles = await CollectResourcesAsync(source, chapters, publication, manuscript, options, diagnostics, cancellationToken).ConfigureAwait(false);
        var styles = new List<string>();
        if (options.IncludeDefaultStyles) {
            publication.AddStylesheet("manuscript-defaults", "EPUB/styles/defaults.css", DefaultCss);
            styles.Add("manuscript-defaults");
        }
        styles.AddRange(sourceStyles.Select(stylesheet => stylesheet.Id));
        for (int index = 0; index < chapters.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            var chapter = chapters[index];
            string id = "chapter-" + (index + 1);
            publication.AddChapter(id, chapter.Path, chapter.Title,
                string.Concat(chapter.Body.Nodes().Select(node => node.ToString(SaveOptions.DisableFormatting))), styles);
            XDocument content = publication.GetContentXml(id);
            content.Root!.Element(Xhtml + "body")!.ReplaceAttributes(chapter.Body.Attributes());
            foreach (string attribute in new[] { "class", "dir" })
                content.Root.SetAttributeValue(attribute, source.DocumentElement?.GetAttribute(attribute));
            XElement[] stylesheetLinks = content.Root.Element(Xhtml + "head")!.Elements(Xhtml + "link").ToArray();
            int offset = options.IncludeDefaultStyles ? 1 : 0;
            for (int styleIndex = 0; styleIndex < sourceStyles.Count; styleIndex++)
                foreach (var attribute in sourceStyles[styleIndex].Attributes) stylesheetLinks[styleIndex + offset].SetAttributeValue(attribute.Key, attribute.Value);
            publication.SetContentXml(id, content);
        }
        publication.SetNavigation(BuildNavigation(chapters));
        try { publication.Write(cancellationToken: cancellationToken); }
        catch (Exception error) when (error is InvalidDataException || error is NotSupportedException || error is XmlException) {
            AddDiagnostic(diagnostics, "EPUB_IMPORT_PACKAGE_INVALID", error.Message, null, OfficeConversionLossKind.Failure);
        }
        return new EpubManuscriptResult(publication, diagnostics);
    }

    private static void AddDiagnostic(List<OfficeConversionFidelityDiagnostic> diagnostics, string code, string message,
        string? location, OfficeConversionLossKind loss = OfficeConversionLossKind.Omission) {
        if (diagnostics.Count >= EpubManuscriptReport.MaximumDiagnostics) return;
        diagnostics.Add(diagnostics.Count == EpubManuscriptReport.MaximumDiagnostics - 1 ? EpubManuscriptReport.LimitDiagnostic() :
            new OfficeConversionFidelityDiagnostic(code, message, loss, "OfficeIMO.Epub.Import", location));
    }
}
