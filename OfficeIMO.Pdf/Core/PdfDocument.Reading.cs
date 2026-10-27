using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Gets metadata through the document's bounded read pipeline.</summary>
    public PdfMetadata Metadata => GetReadSnapshot().Document.Metadata;
    /// <summary>Gets document outlines without exposing the parser.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines => GetReadSnapshot().Document.Outlines;
    /// <summary>Gets named destinations from the document.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations => GetReadSnapshot().Document.NamedDestinations;
    /// <summary>Gets the document's form fields.</summary>
    public IReadOnlyList<PdfFormField> FormFields => GetReadSnapshot().Document.FormFields;
    /// <summary>Gets structural recovery observations from the bounded parser.</summary>
    public PdfRepairReport RepairReport => GetReadSnapshot().Document.RepairReport;

    /// <summary>Extracts page text through the document's bounded read pipeline.</summary>
    public string ExtractText(CancellationToken cancellationToken = default) =>
        GetReadSnapshot(cancellationToken: cancellationToken).Document.ExtractText(cancellationToken);

    /// <summary>Extracts Markdown text from the bounded logical reading model.</summary>
    public string ExtractMarkdown(PdfLogicalMarkdownOptions? options = null, CancellationToken cancellationToken = default) =>
        Read(cancellationToken: cancellationToken).ExtractMarkdown(options);

    /// <summary>Extracts text with column-aware reading order.</summary>
    public string ExtractTextWithColumns(PdfTextLayoutOptions? options = null) =>
        GetReadSnapshot().Document.ExtractTextWithColumns(options);

    /// <summary>Extracts structured pages with paragraphs, lists, headings and tables.</summary>
    public IReadOnlyList<PdfStructuredPage> ExtractStructuredPages(PdfTextLayoutOptions? options = null) =>
        GetReadSnapshot().Document.ExtractStructuredPages(options);

    /// <summary>Extracts detected tables grouped by page.</summary>
    public IReadOnlyList<PdfStructuredTablePage> ExtractTablesByPage(PdfTextLayoutOptions? options = null) =>
        GetReadSnapshot().Document.ExtractTablesByPage(options);

    /// <summary>Extracts paragraphs grouped by page.</summary>
    public IReadOnlyList<PdfStructuredParagraphPage> ExtractParagraphsByPage(PdfTextLayoutOptions? options = null) =>
        GetReadSnapshot().Document.ExtractParagraphsByPage(options);

    /// <summary>Extracts detected headings grouped by page.</summary>
    public IReadOnlyList<PdfStructuredHeadingPage> ExtractHeadingsByPage(PdfTextLayoutOptions? options = null) =>
        GetReadSnapshot().Document.ExtractHeadingsByPage(options);

    /// <summary>Extracts list items grouped by page.</summary>
    public IReadOnlyList<PdfStructuredListItemPage> ExtractListItemsByPage(PdfTextLayoutOptions? options = null) =>
        GetReadSnapshot().Document.ExtractListItemsByPage(options);

    /// <summary>Gets the parsed document's tagged content.</summary>
    public PdfTaggedContentInfo? TaggedContent => GetReadSnapshot().Document.TaggedContent;

    /// <summary>Gets the parsed document's has tagged content.</summary>
    public bool HasTaggedContent => GetReadSnapshot().Document.HasTaggedContent;

    /// <summary>Gets the parsed document's catalog actions.</summary>
    public IReadOnlyList<PdfCatalogAction> CatalogActions => GetReadSnapshot().Document.CatalogActions;

    /// <summary>Gets the parsed document's java scripts.</summary>
    public IReadOnlyList<PdfJavaScript> JavaScripts => GetReadSnapshot().Document.JavaScripts;

    /// <summary>Gets the parsed document's portfolio.</summary>
    public PdfPortfolioInfo? Portfolio => GetReadSnapshot().Document.Portfolio;

    /// <summary>Gets the parsed document's acro form default appearance.</summary>
    public string? AcroFormDefaultAppearance => GetReadSnapshot().Document.AcroFormDefaultAppearance;

    /// <summary>Gets the parsed document's acro form quadding.</summary>
    public int? AcroFormQuadding => GetReadSnapshot().Document.AcroFormQuadding;

    /// <summary>Gets the parsed document's acro form need appearances.</summary>
    public bool? AcroFormNeedAppearances => GetReadSnapshot().Document.AcroFormNeedAppearances;

    /// <summary>Gets the parsed document's acro form signature flags.</summary>
    public int? AcroFormSignatureFlags => GetReadSnapshot().Document.AcroFormSignatureFlags;

    /// <summary>Gets the parsed document's acro form xfa.</summary>
    public PdfAcroFormXfaInfo? AcroFormXfa => GetReadSnapshot().Document.AcroFormXfa;

    /// <summary>Gets the parsed document's catalog version.</summary>
    public string? CatalogVersion => GetReadSnapshot().Document.CatalogVersion;

    /// <summary>Gets the parsed document's catalog language.</summary>
    public string? CatalogLanguage => GetReadSnapshot().Document.CatalogLanguage;

    /// <summary>Gets the parsed document's output intents.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntents => GetReadSnapshot().Document.OutputIntents;

    /// <summary>Gets the parsed document's output intents are complete.</summary>
    public bool OutputIntentsAreComplete => GetReadSnapshot().Document.OutputIntentsAreComplete;

    /// <summary>Gets the parsed document's optional content.</summary>
    public PdfOptionalContentProperties? OptionalContent => GetReadSnapshot().Document.OptionalContent;

    /// <summary>Gets the parsed document's xmp metadata.</summary>
    public PdfXmpMetadataInfo? XmpMetadata => GetReadSnapshot().Document.XmpMetadata;

    /// <summary>Extracts image XObjects from all pages in document order.</summary>
    public IReadOnlyList<PdfExtractedImage> ExtractImages() => GetReadSnapshot().Document.ExtractImages();

    /// <summary>Extracts the document's embedded attachment payloads.</summary>
    public IReadOnlyList<PdfExtractedAttachment> ExtractAttachments() => GetReadSnapshot().Document.ExtractAttachments();
}
