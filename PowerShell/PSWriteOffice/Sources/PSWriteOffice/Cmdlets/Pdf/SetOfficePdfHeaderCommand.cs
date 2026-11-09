using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Composes running PDF Header text, child runs or the native builder inside New-OfficePdf.</summary>
/// <example><summary>Add a styled running Header.</summary>
/// <code>New-OfficePdf ./Report.pdf { PdfHeader { PdfText 'Internal' -Italic }; PdfParagraph 'Report' }</code></example>
[Cmdlet(VerbsCommon.Set, "OfficePdfHeader", DefaultParameterSetName = "Text")]
[OutputType(typeof(void))]
[Alias("PdfHeader") ]
public sealed class SetOfficePdfHeaderCommand : PSCmdlet {
    /// <summary>Running text supporting {page} and {pages} tokens.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Styled PdfText child commands.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Advanced native Header builder, passed as the script argument.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Builder")]
    public ScriptBlock? Compose { get; set; }
    /// <summary>Text alignment.</summary>
    [Parameter(ParameterSetName = "Text"), Parameter(ParameterSetName = "Content")]
    public OfficeAlignment Alignment { get; set; } = OfficeAlignment.Center;
    /// <summary>Text size; bare numbers mean points.</summary>
    [Parameter(ParameterSetName = "Text"), Parameter(ParameterSetName = "Content"), OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Configure the first page variant.</summary>
    [Parameter(ParameterSetName = "Text"), Parameter(ParameterSetName = "Content")]
    public SwitchParameter FirstPage { get; set; }
    /// <summary>Configure the even page variant.</summary>
    [Parameter(ParameterSetName = "Text"), Parameter(ParameterSetName = "Content")]
    public SwitchParameter EvenPage { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var runs = new PdfRunningTextContent();
        if (Content != null) {
            using (OfficeDocumentContext.Push(runs)) Content.InvokeReturnAsIs();
        }
        var compose = Compose?.GetNewClosure();
        string text = Text;
        PdfDslContext.Require(this).ConfigurePage(page => page.Header(builder => {
            if (compose != null) { compose.Invoke(builder); return; }
            builder.Alignment(Alignment);
            if (Size.HasValue) builder.FontSize(Size.Value);
            if (Content == null) {
                if (FirstPage) builder.FirstPageText(text);
                if (EvenPage) builder.EvenPagesText(text);
                if (!FirstPage && !EvenPage) builder.Text(text);
            } else {
                System.Action<PdfHeaderTextBuilder> add = runs.Apply;
                if (FirstPage) builder.FirstPageText(add);
                if (EvenPage) builder.EvenPagesText(add);
                if (!FirstPage && !EvenPage) builder.Text(add);
            }
        }));
    }
}
