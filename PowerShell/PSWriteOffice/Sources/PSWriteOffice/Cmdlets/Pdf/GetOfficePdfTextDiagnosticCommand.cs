using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets PDF text encoding and advanced-layout diagnostics for generated text before rendering.</summary>
/// <example>
///   <summary>Check text that needs complex shaping.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePdfTextDiagnostic -Text 'مرحبا' -AdvancedLayout</code>
///   <para>Returns OfficeIMO.Pdf diagnostics describing right-to-left and complex-script layout requirements.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfTextDiagnostic")]
[OutputType(typeof(PdfTextEncodingDiagnostic), typeof(PdfTextShapingDiagnostic))]
public sealed class GetOfficePdfTextDiagnosticCommand : PSCmdlet
{
    /// <summary>Text to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Optional source label included in diagnostic objects.</summary>
    [Parameter]
    public string Source { get; set; } = string.Empty;

    /// <summary>Optional TrueType or OpenType/CFF font used for embedded-font glyph coverage and layout diagnostics.</summary>
    [Parameter]
    public string? FontPath { get; set; }

    /// <summary>Emit only encoding/glyph coverage diagnostics.</summary>
    [Parameter]
    public SwitchParameter Encoding { get; set; }

    /// <summary>Emit only advanced-layout diagnostics such as right-to-left, complex-script shaping, mark positioning, and script line breaking.</summary>
    [Parameter]
    public SwitchParameter AdvancedLayout { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        bool includeEncoding = Encoding.IsPresent || !AdvancedLayout.IsPresent;
        bool includeAdvanced = AdvancedLayout.IsPresent || !Encoding.IsPresent;

        byte[]? fontBytes = null;
        string? fontName = null;
        if (!string.IsNullOrWhiteSpace(FontPath))
        {
            var resolvedFontPath = PdfCommandUtilities.ResolvePath(this, FontPath!);
            fontBytes = File.ReadAllBytes(resolvedFontPath);
            fontName = Path.GetFileNameWithoutExtension(resolvedFontPath);
        }

        if (includeEncoding)
        {
            var encodingDiagnostics = fontBytes == null
                ? PdfTextPreflight.AnalyzeWinAnsi(Text, Source)
                : PdfTextPreflight.AnalyzeEmbeddedFont(Text, fontBytes, Source, fontName);
            WriteObject(encodingDiagnostics, enumerateCollection: true);
        }

        if (includeAdvanced)
        {
            var shapingDiagnostics = fontBytes == null
                ? PdfTextPreflight.AnalyzeAdvancedLayout(Text, Source)
                : PdfTextPreflight.AnalyzeAdvancedLayout(Text, fontBytes, Source, fontName);
            WriteObject(shapingDiagnostics, enumerateCollection: true);
        }
    }
}
