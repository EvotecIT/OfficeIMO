using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates discoverable page and rendering settings for Export-OfficeWordImage.</summary>
/// <example>
///   <summary>Render the first two pages at higher density.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordImageOptions -PageIndex 0 -PageCount 2 -TargetDpi 144 -IncludeDocumentContent
/// Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Pages -Options $options</code>
///   <para>Supplying <c>PageCount</c> selects batch export, so <c>OutputPath</c> is a folder. Use <c>-AllPages</c> on the export command for the complete document.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordImageOptions")]
[OutputType(typeof(WordImageExportOptions))]
public sealed class NewOfficeWordImageOptionsCommand : OfficeImageOptionsCommandBase<WordImageExportOptions> {
    /// <summary>Render document content.</summary>
    [Parameter] public SwitchParameter IncludeDocumentContent { get; set; }
    /// <summary>Zero-based first page index.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? PageIndex { get; set; }
    /// <summary>Maximum pages exported. Supplying this value selects batch export.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? PageCount { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordImageExportOptions();
        ApplyCommon(options);
        if (IsBound(nameof(IncludeDocumentContent))) options.IncludeDocumentContent = IncludeDocumentContent.IsPresent;
        if (PageIndex.HasValue) options.PageIndex = PageIndex.Value;
        if (PageCount.HasValue) options.PageCount = PageCount.Value;
        WriteObject(options);
    }
}
