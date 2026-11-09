using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Permanently removes (redacts) text, pattern matches or areas from a PDF and writes a new PDF.</summary>
/// <example><summary>Use the native OfficePdfContent command.</summary><code>Remove-OfficePdfContent -Path .\Contracts\MSA-Contoso.pdf -OutputPath .\Contracts\MSA-Redacted.pdf -Text 'Contoso Ltd'</code></example>
[Cmdlet("Remove", "OfficePdfContent", DefaultParameterSetName = "Text", SupportsShouldProcess = true)]
[OutputType(typeof(object))]
public sealed class RemoveOfficePdfContentCommand : OfficeCmdlet {
    private readonly List<string> _inputs = new();
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "TextDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Pattern", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PatternDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Area", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "AreaDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true, Position = 0)]
    public string[]? Path { get; set; }
    /// <summary>OutputPath for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Pattern", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Area", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>Text for this operation.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "TextDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory", Mandatory = true)]
    public string[]? Text { get; set; }
    /// <summary>Replaces -MatchCase.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    public SwitchParameter CaseSensitive { get; set; }
    /// <summary>FillColor for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaDirectory")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }
    /// <summary>OnlyPaintMatches for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    public SwitchParameter OnlyPaintMatches { get; set; }
    /// <summary>Argument transformation also accepts a plain string (with a warning) or a PSCredential (its password).</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaDirectory")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    [OfficePasswordTransformation]
    public System.Security.SecureString? Password { get; set; }
    /// <summary>After successful authentication only; never bypasses a missing password.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaDirectory")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>LoadOptions for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaDirectory")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>NoClobber for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaDirectory")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextDirectory")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternDirectory")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaDirectory")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>OutputDirectory for this operation.</summary>
    [Parameter(ParameterSetName = "TextDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>LiteralPath for this operation.</summary>
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "TextLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "PatternLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "AreaLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[]? LiteralPath { get; set; }
    /// <summary>Bounded regular expressions matched with the same case policy as Get-OfficePdfRedactionPlan.</summary>
    [Parameter(ParameterSetName = "Pattern", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternLiteralPathDirectory", Mandatory = true)]
    public string[]? Pattern { get; set; }
    /// <summary>Area for this operation.</summary>
    [Parameter(ParameterSetName = "Area", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaLiteralPathDirectory", Mandatory = true)]
    public PdfRedactionArea[]? Area { get; set; }
    /// <summary>1-based.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true)]
    public int? Page { get; set; }
    /// <summary>X for this operation.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength? X { get; set; }
    /// <summary>Y for this operation.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength? Y { get; set; }
    /// <summary>Width for this operation.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }
    /// <summary>Height for this operation.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }
    /// <summary>Label for this operation.</summary>
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleDirectory")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [Parameter(ParameterSetName = "RectangleLiteralPathDirectory")]
    public string? Label { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        _inputs.AddRange(OfficePath.ResolveInputs(this, LiteralPath ?? Path ?? Array.Empty<string>(), LiteralPath != null));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        foreach (var pair in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, NoClobber)) {
            if (!ShouldProcess(pair.Output, "Remove PDF content permanently")) continue;
            using var document = PdfDocument.Load(pair.Input, PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions));
            var options = new PdfRedactionApplyOptions { FillColor = FillColor ?? OfficeColor.Black, PaintUnmatchedAreas = !OnlyPaintMatches };
            PdfDocument result;
            if (Text != null || Pattern != null) {
                var search = PdfRedactionCommandOptions.Create(Text, Pattern, CaseSensitive.IsPresent);
                var plan = document.Redactions.Search(search);
                if (!plan.HasMatches) throw new PSInvalidOperationException("No PDF text matched. The output was not written.");
                result = document.Redactions.Apply(plan, options);
            } else result = document.Redactions.Apply(Area ?? new[] { new PdfRedactionArea(Page!.Value, X!.Value.ToPoints(), Y!.Value.ToPoints(), Width!.Value.ToPoints(), Height!.Value.ToPoints(), Label) }, options);
            PdfCommandUtilities.EnsureDirectory(pair.Output);
            bool transferred = false;
            try {
                result.Save(pair.Output, new PdfSaveOptions { ConflictPolicy = NoClobber
                    ? OfficeConversionFileConflictPolicy.FailIfExists : OfficeConversionFileConflictPolicy.Replace });
                WriteObject(PassThru ? (object)result : new FileInfo(pair.Output));
                transferred = PassThru.IsPresent;
            } finally { if (!transferred) result.Dispose(); }
        }
    }
}
