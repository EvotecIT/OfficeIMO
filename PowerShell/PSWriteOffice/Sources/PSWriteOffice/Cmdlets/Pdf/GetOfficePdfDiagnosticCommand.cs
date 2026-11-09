using OfficeIMO;
using System.Management.Automation;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Returns native PDF structure, preflight, incremental, optimization or generated-text diagnostics.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdfDiagnostic", DefaultParameterSetName = "Path")]
public sealed class GetOfficePdfDiagnosticCommand : PSCmdlet {
    /// <summary>Input PDF files with wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path"), SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal input files.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Native report to return for each PDF.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    public OfficePdfDiagnosticType Type { get; set; } = OfficePdfDiagnosticType.Structure;
    /// <summary>Source password.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath"), OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>After authentication, explicitly ignore owner-imposed restrictions.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native source loading policy.</summary>
    [Parameter(ParameterSetName = "Path"), Parameter(ParameterSetName = "LiteralPath"), OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Generated text to diagnose before rendering.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Source label included in text findings.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string Source { get; set; } = string.Empty;
    /// <summary>Literal font file for glyph coverage and shaping diagnostics.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? FontPath { get; set; }
    /// <summary>Return encoding and glyph coverage findings.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter Encoding { get; set; }
    /// <summary>Return script, direction, mark positioning and line-breaking findings.</summary>
    [Parameter(ParameterSetName = "Text")]
    public SwitchParameter AdvancedLayout { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Text") { DiagnoseText(); return; }
        var read = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions.IsPresent);
        if (OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(IgnorePermissionRestrictions)) is bool ignore)
            read = (read ?? new PdfLoadOptions()).Snapshot(null, null, ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce);
        foreach (var path in OfficePath.ResolveInputs(this, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath")) {
            // Preflight must remain reachable for documents whose full read is blocked.
            if (Type == OfficePdfDiagnosticType.Preflight) { WriteObject(OfficeDocument.Inspect(path, new OfficeInspectOptions { IncludeContent = false, PdfPreflight = read ?? new PdfLoadOptions() }).PdfPreflight); continue; }
            using var document = PdfDocument.Load(path, read);
            WriteObject(Type switch {
                OfficePdfDiagnosticType.Structure => (object)document.Diagnostics(),
                OfficePdfDiagnosticType.Incremental => document.AnalyzeAppendOnlyMutation(),
                OfficePdfDiagnosticType.Optimization => document.Optimization.Analyze(),
                _ => throw new PSArgumentOutOfRangeException(nameof(Type), Type, "Unknown PDF diagnostic report.")
            });
        }
    }
    private void DiagnoseText() {
        bool includeEncoding = Encoding.IsPresent || !AdvancedLayout.IsPresent;
        bool includeAdvanced = AdvancedLayout.IsPresent || !Encoding.IsPresent;

        byte[]? fontBytes = null;
        string? fontName = null;
        if (!string.IsNullOrWhiteSpace(FontPath))
        {
            var resolvedFontPath = PdfCommandUtilities.ResolvePath(this, FontPath!);
            fontBytes = File.ReadAllBytes(resolvedFontPath);
            fontName = System.IO.Path.GetFileNameWithoutExtension(resolvedFontPath);
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
