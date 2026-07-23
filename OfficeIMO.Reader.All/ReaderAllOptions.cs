namespace OfficeIMO.Reader.All;

/// <summary>
/// Optional format-specific settings used by <see cref="OfficeDocumentReaderBuilderAllExtensions.AddAllOfficeIMOHandlers"/>.
/// </summary>
/// <remarks>
/// Each adapter captures a defensive copy during registration. Changing this object after the preset is applied
/// does not change a built reader.
/// </remarks>
public sealed class ReaderAllOptions {
    /// <summary>Gets or sets AsciiDoc adapter options.</summary>
    public AsciiDoc.ReaderAsciiDocOptions? AsciiDoc { get; set; }

    /// <summary>Gets or sets CSV and TSV adapter options.</summary>
    public Csv.CsvReadOptions? Csv { get; set; }

    /// <summary>Gets or sets EPUB adapter options.</summary>
    public OfficeIMO.Epub.EpubReadOptions? Epub { get; set; }

    /// <summary>Gets or sets direct email, store, and Offline Address Book adapter options.</summary>
    public Email.ReaderEmailHandlersOptions? Email { get; set; }

    /// <summary>Gets or sets Excel adapter options.</summary>
    public Excel.ReaderExcelOptions? Excel { get; set; }

    /// <summary>Gets or sets HTML adapter options.</summary>
    public Html.ReaderHtmlOptions? Html { get; set; }

    /// <summary>Gets or sets standalone image adapter options.</summary>
    public Image.ReaderImageOptions? Image { get; set; }

    /// <summary>Gets or sets JSON adapter options.</summary>
    public Json.JsonReadOptions? Json { get; set; }

    /// <summary>Gets or sets LaTeX adapter options.</summary>
    public Latex.ReaderLatexOptions? Latex { get; set; }

    /// <summary>Gets or sets Markdown adapter options.</summary>
    public Markdown.ReaderMarkdownOptions? Markdown { get; set; }

    /// <summary>Gets or sets Jupyter Notebook adapter options.</summary>
    public Notebook.ReaderNotebookOptions? Notebook { get; set; }

    /// <summary>
    /// Gets or sets offline OneNote adapter options. When omitted, the all-formats preset does not follow sibling
    /// files referenced by a <c>.onetoc2</c> path. Supply explicit options to opt into that compatibility behavior.
    /// </summary>
    public OneNote.ReaderOneNoteOptions? OneNote { get; set; }

    /// <summary>Gets or sets OpenDocument adapter options.</summary>
    public OpenDocument.ReaderOpenDocumentOptions? OpenDocument { get; set; }

    /// <summary>Gets or sets PDF adapter options.</summary>
    public Pdf.ReaderPdfOptions? Pdf { get; set; }

    /// <summary>Gets or sets PowerPoint adapter options.</summary>
    public PowerPoint.ReaderPowerPointOptions? PowerPoint { get; set; }

    /// <summary>Gets or sets RTF adapter options.</summary>
    public Rtf.ReaderRtfOptions? Rtf { get; set; }

    /// <summary>Gets or sets SRT and WebVTT adapter options.</summary>
    public Subtitles.ReaderSubtitleOptions? Subtitles { get; set; }

    /// <summary>Gets or sets Visio adapter options.</summary>
    public Visio.ReaderVisioOptions? Visio { get; set; }

    /// <summary>Gets or sets Word adapter options.</summary>
    public Word.ReaderWordOptions? Word { get; set; }

    /// <summary>Gets or sets XML adapter options.</summary>
    public Xml.XmlReadOptions? Xml { get; set; }

    /// <summary>Gets or sets YAML adapter options.</summary>
    public Yaml.YamlReadOptions? Yaml { get; set; }

    /// <summary>Gets or sets ZIP archive traversal limits.</summary>
    public OfficeIMO.Zip.ZipTraversalOptions? ZipTraversal { get; set; }

    /// <summary>Gets or sets Reader-specific nested ZIP behavior.</summary>
    public Zip.ReaderZipOptions? Zip { get; set; }
}
