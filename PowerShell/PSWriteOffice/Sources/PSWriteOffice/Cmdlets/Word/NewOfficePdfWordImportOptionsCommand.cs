using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates discoverable PDF-to-Word reconstruction settings.</summary>
/// <example>
///   <summary>Reconstruct headings, paragraphs, lists, and tables.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePdfToWordOptions -ImportHeadings -ImportParagraphs -ImportLists -ImportTables
/// ConvertTo-OfficePdfWord -Path .\Source.pdf -OutputPath .\Rebuilt.docx -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfToWordOptions")]
[Alias("New-OfficePdfWordImportOptions")]
[OutputType(typeof(PdfToWordOptions))]
public sealed class NewOfficePdfToWordOptionsCommand : PSCmdlet {
    /// <summary>Use the built-in tables-only import profile.</summary>
    [Parameter] public SwitchParameter TablesOnly { get; set; }
    /// <summary>Copy PDF metadata into Word properties.</summary>
    [Parameter] public SwitchParameter IncludeMetadata { get; set; }
    /// <summary>Represent source pages with Word page breaks.</summary>
    [Parameter] public SwitchParameter PreservePageBreaks { get; set; }
    /// <summary>Represent empty PDF pages.</summary>
    [Parameter] public SwitchParameter IncludeEmptyPages { get; set; }
    /// <summary>Import detected headings.</summary>
    [Parameter] public SwitchParameter ImportHeadings { get; set; }
    /// <summary>Import detected paragraphs.</summary>
    [Parameter] public SwitchParameter ImportParagraphs { get; set; }
    /// <summary>Use the crop-, rotation-, and column-aware reading order.</summary>
    [Parameter] public SwitchParameter UseSharedPageReadingOrder { get; set; }
    /// <summary>Import detected lists.</summary>
    [Parameter] public SwitchParameter ImportLists { get; set; }
    /// <summary>Import detected tables.</summary>
    [Parameter] public SwitchParameter ImportTables { get; set; }
    /// <summary>Import safe URI links.</summary>
    [Parameter] public SwitchParameter ImportUriLinks { get; set; }
    /// <summary>Import supported internal links.</summary>
    [Parameter] public SwitchParameter ImportInternalLinks { get; set; }
    /// <summary>Prefix for generated Word bookmarks.</summary>
    [Parameter] public string? BookmarkPrefix { get; set; }
    /// <summary>Allowed absolute hyperlink URI schemes.</summary>
    [Parameter] public string[]? AllowedHyperlinkUriScheme { get; set; }
    /// <summary>Import supported embedded images.</summary>
    [Parameter] public SwitchParameter ImportImages { get; set; }
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
        PdfToWordOptions options = TablesOnly.IsPresent ? PdfToWordOptions.CreateTablesOnly() : new PdfToWordOptions();
        Apply(nameof(IncludeMetadata), value => options.IncludeMetadata = value);
        Apply(nameof(PreservePageBreaks), value => options.PreservePageBreaks = value);
        Apply(nameof(IncludeEmptyPages), value => options.IncludeEmptyPages = value);
        Apply(nameof(ImportHeadings), value => options.ImportHeadings = value);
        Apply(nameof(ImportParagraphs), value => options.ImportParagraphs = value);
        Apply(nameof(UseSharedPageReadingOrder), value => options.UseSharedPageReadingOrder = value);
        Apply(nameof(ImportLists), value => options.ImportLists = value);
        Apply(nameof(ImportTables), value => options.ImportTables = value);
        Apply(nameof(ImportUriLinks), value => options.ImportUriLinks = value);
        Apply(nameof(ImportInternalLinks), value => options.ImportInternalLinks = value);
        Apply(nameof(ImportImages), value => options.ImportImages = value);
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
        if (!MyInvocation.BoundParameters.ContainsKey(name)) return;
        setter(((SwitchParameter)MyInvocation.BoundParameters[name]).IsPresent);
    }
}
