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

    /// <summary>Gets or sets PST, OST, OLM, and EMLX adapter options.</summary>
    public EmailStore.ReaderEmailStoreOptions? EmailStore { get; set; }

    /// <summary>Gets or sets Outlook Offline Address Book adapter options.</summary>
    public EmailAddressBook.ReaderEmailAddressBookOptions? EmailAddressBook { get; set; }

    /// <summary>Gets or sets HTML adapter options.</summary>
    public Html.ReaderHtmlOptions? Html { get; set; }

    /// <summary>Gets or sets standalone image adapter options.</summary>
    public Image.ReaderImageOptions? Image { get; set; }

    /// <summary>Gets or sets JSON adapter options.</summary>
    public Json.JsonReadOptions? Json { get; set; }

    /// <summary>Gets or sets LaTeX adapter options.</summary>
    public Latex.ReaderLatexOptions? Latex { get; set; }

    /// <summary>Gets or sets Jupyter Notebook adapter options.</summary>
    public Notebook.ReaderNotebookOptions? Notebook { get; set; }

    /// <summary>Gets or sets offline OneNote adapter options.</summary>
    public OneNote.ReaderOneNoteOptions? OneNote { get; set; }

    /// <summary>Gets or sets PDF adapter options.</summary>
    public Pdf.ReaderPdfOptions? Pdf { get; set; }

    /// <summary>Gets or sets RTF adapter options.</summary>
    public Rtf.ReaderRtfOptions? Rtf { get; set; }

    /// <summary>Gets or sets SRT and WebVTT adapter options.</summary>
    public Subtitles.ReaderSubtitleOptions? Subtitles { get; set; }

    /// <summary>Gets or sets Visio adapter options.</summary>
    public Visio.ReaderVisioOptions? Visio { get; set; }

    /// <summary>Gets or sets XML adapter options.</summary>
    public Xml.XmlReadOptions? Xml { get; set; }

    /// <summary>Gets or sets YAML adapter options.</summary>
    public Yaml.YamlReadOptions? Yaml { get; set; }

    /// <summary>Gets or sets ZIP archive traversal limits.</summary>
    public OfficeIMO.Zip.ZipTraversalOptions? ZipTraversal { get; set; }

    /// <summary>Gets or sets Reader-specific nested ZIP behavior.</summary>
    public Zip.ReaderZipOptions? Zip { get; set; }
}
