using OfficeIMO.Pdf.Word;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates discoverable PDF-to-Word reconstruction settings.</summary>
/// <example><summary>Reconstruct editable Word content while omitting images.</summary>
/// <code>$options = New-OfficePdfWordOptions -NoImages
/// Convert-OfficeDocument -Path ./Source.pdf -To Word -OutputPath ./Rebuilt.docx -Options $options</code></example>
/// <example><summary>Import tables without the surrounding paragraphs.</summary>
/// <code>$options = New-OfficePdfWordOptions -TablesOnly
/// Convert-OfficeDocument -Path ./Tables.pdf -To Word -OutputPath ./Tables.docx -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfWordOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordPdfImportOptions))]
public sealed class NewOfficePdfWordOptionsCommand : OfficeCmdlet {
    /// <summary>Use the built-in tables-only import profile.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter TablesOnly { get; set; }
    /// <summary>Copy PDF metadata into Word properties.</summary>
    [Parameter] public SwitchParameter IncludeMetadata { get; set; }
    /// <summary>Represent source pages with Word page breaks.</summary>
    [Parameter] public SwitchParameter PreservePageBreaks { get; set; }
    /// <summary>Represent empty PDF pages.</summary>
    [Parameter] public SwitchParameter IncludeEmptyPages { get; set; }
    /// <summary>Exclude detected headings.</summary>
    [Parameter] public SwitchParameter NoHeadings { get; set; }
    /// <summary>Exclude detected paragraphs.</summary>
    [Parameter] public SwitchParameter NoParagraphs { get; set; }
    /// <summary>Use the crop-, rotation-, and column-aware reading order.</summary>
    [Parameter] public SwitchParameter UseSharedPageReadingOrder { get; set; }
    /// <summary>Exclude detected lists.</summary>
    [Parameter] public SwitchParameter NoLists { get; set; }
    /// <summary>Exclude detected tables.</summary>
    [Parameter] public SwitchParameter NoTables { get; set; }
    /// <summary>Exclude safe URI links.</summary>
    [Parameter] public SwitchParameter NoUriLinks { get; set; }
    /// <summary>Exclude supported internal links.</summary>
    [Parameter] public SwitchParameter NoInternalLinks { get; set; }
    /// <summary>Prefix for generated Word bookmarks.</summary>
    [Parameter] public string? BookmarkPrefix { get; set; }
    /// <summary>Allowed absolute hyperlink URI schemes.</summary>
    [Parameter] public string[]? AllowedHyperlinkUriScheme { get; set; }
    /// <summary>Exclude supported embedded images.</summary>
    [Parameter] public SwitchParameter NoImages { get; set; }
    /// <summary>Preserve detected image placement size.</summary>
    [Parameter] public SwitchParameter PreserveImagePlacementSize { get; set; }
    /// <summary>Use paragraphs when an image cannot be embedded.</summary>
    [Parameter] public SwitchParameter IncludeImagePlaceholders { get; set; }
    /// <summary>Represent AcroForm widgets with editable placeholders.</summary>
    [Parameter] public SwitchParameter IncludeFormFieldPlaceholders { get; set; }
    /// <summary>Maximum body rows imported per table; zero means unlimited.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaxTableRows { get; set; }
    /// <summary>Word table style for imported tables.</summary>
    [Parameter] public WordTableStyle? TableStyle { get; set; }
    /// <summary>Repeat inferred table header rows.</summary>
    [Parameter] public SwitchParameter RepeatHeaderRows { get; set; }
    /// <summary>Fit imported tables to page width.</summary>
    [Parameter] public SwitchParameter FitTablesToPageWidth { get; set; }
    /// <summary>Right-align inferred numeric columns.</summary>
    [Parameter] public SwitchParameter AlignNumericColumns { get; set; }
    /// <summary>Text used when no supported content is detected.</summary>
    [Parameter] public string? EmptyDocumentMessage { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        WordPdfImportOptions options = TablesOnly.IsPresent ? WordPdfImportOptions.CreateTablesOnly() : new WordPdfImportOptions();
        Apply(nameof(IncludeMetadata), value => options.IncludeMetadata = value);
        Apply(nameof(PreservePageBreaks), value => options.PreservePageBreaks = value);
        Apply(nameof(IncludeEmptyPages), value => options.IncludeEmptyPages = value);
        Apply(nameof(NoHeadings), value => options.ImportHeadings = !value);
        Apply(nameof(NoParagraphs), value => options.ImportParagraphs = !value);
        Apply(nameof(UseSharedPageReadingOrder), value => options.UseSharedPageReadingOrder = value);
        Apply(nameof(NoLists), value => options.ImportLists = !value);
        Apply(nameof(NoTables), value => options.ImportTables = !value);
        Apply(nameof(NoUriLinks), value => options.ImportUriLinks = !value);
        Apply(nameof(NoInternalLinks), value => options.ImportInternalLinks = !value);
        Apply(nameof(NoImages), value => options.ImportImages = !value);
        Apply(nameof(PreserveImagePlacementSize), value => options.PreserveImagePlacementSize = value);
        Apply(nameof(IncludeImagePlaceholders), value => options.IncludeImagePlaceholders = value);
        Apply(nameof(IncludeFormFieldPlaceholders), value => options.IncludeFormFieldPlaceholders = value);
        Apply(nameof(RepeatHeaderRows), value => options.RepeatHeaderRows = value);
        Apply(nameof(FitTablesToPageWidth), value => options.FitTablesToPageWidth = value);
        Apply(nameof(AlignNumericColumns), value => options.AlignNumericColumns = value);
        if (!string.IsNullOrWhiteSpace(BookmarkPrefix)) options.BookmarkPrefix = BookmarkPrefix!;
        if (AllowedHyperlinkUriScheme != null) {
            options.AllowedHyperlinkUriSchemes.Clear();
            foreach (string scheme in AllowedHyperlinkUriScheme) if (!string.IsNullOrWhiteSpace(scheme)) options.AllowedHyperlinkUriSchemes.Add(scheme);
        }
        if (MaxTableRows.HasValue) options.MaxTableRows = MaxTableRows.Value;
        if (TableStyle.HasValue) options.TableStyle = TableStyle.Value;
        if (EmptyDocumentMessage != null) options.EmptyDocumentMessage = EmptyDocumentMessage;
        WriteObject(options);
    }

    private void Apply(string name, System.Action<bool> setter) {
        if (BoundSwitch(name) is bool value) setter(value);
    }
}
