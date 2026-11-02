using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Opens native PDF documents or atomically saves a completed editing block.</summary>
/// <para>Without Content returns live documents without writing files. A block takes one input and saves its final edits to OutputPath or back to the source. Opening preserves the existing theme and security contract.</para>
/// <example><summary>Edit a literal file and save a copy.</summary>
/// <code>Get-Item -LiteralPath './Review[final].pdf' | Open-OfficePdf -OutputPath ./Reviewed.pdf -PassThru -Content {
///     Remove-OfficePdfPage -Page last
///     PdfMetadata -Title 'Reviewed report'
/// }</code></example>
/// <example><summary>Open matching documents for inspection.</summary>
/// <code>$documents = Open-OfficePdf -Path './Reports/*.pdf'
/// try { $documents | ForEach-Object { $_.Read().Text } }
/// finally { $documents | Close-OfficePdf }</code></example>
[Cmdlet(VerbsCommon.Open, "OfficePdf", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class OpenOfficePdfCommand : OfficeLifecycleCmdlet<PdfDocument> {
    private string? _contentInput;
    private int _contentInputCount;
    /// <summary>One or more PDF input paths, with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal PDF paths; accepts FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>Optional destination for a content block; omission saves back to the input.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>Source password as a SecureString, credential, or warned plaintext value.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }

    /// <summary>After authentication, explicitly ignore owner-imposed usage restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Native PDF parsing, resource-limit and authentication settings.</summary>
    [Parameter]
    public PdfLoadOptions? LoadOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Content == null && OutputPath != null) {
            InvalidArgument(new PSArgumentException("OutputPath requires a content block."), OutputPath);
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        var inputs = new List<string>();
        foreach (string path in literal ? LiteralPath : Path)
            inputs.AddRange(OfficePath.ResolveInput(SessionState, path, literal));
        if (Content != null) {
            foreach (string input in inputs) {
                if (_contentInputCount == 0) _contentInput = input;
                _contentInputCount = Math.Min(2, _contentInputCount + 1);
            }
            return;
        }
        foreach (string input in inputs) RunOpened(input);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (Content == null) return;
        if (_contentInputCount != 1) {
            InvalidArgument(new PSArgumentException("A content block requires exactly one input PDF for the entire invocation."), _contentInput);
            return;
        }
        RunOpened(_contentInput!);
    }

    private void RunOpened(string input) {
        var options = PdfCommandUtilities.CreateReadOptions(LoadOptions,
            Password == null ? null : new NetworkCredential(string.Empty, Password).Password,
            IgnorePermissionRestrictions.IsPresent);
        RunDocument(OutputPath ?? input, isNew: false, force: false,
            path => PdfDocument.Load(path!, options),
            (document, path, allowLoss) => {
                document.Save(path, new PdfSaveOptions { AllowLoss = allowLoss });
                return null;
            }, sourcePath: input);
    }
}
