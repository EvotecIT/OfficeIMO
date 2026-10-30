using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Security;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Combines documents of the same native format in input order.</summary>
/// <para>Supports PDF and Open XML Word and Excel files. Binary projection and conversion are not performed.
/// Word and Excel live documents can receive appended files and remain caller-owned. PDF merging creates a new
/// document, so live PDF targets are not supported. PassThru returns a saved FileInfo for file input or the live
/// document for Document input. Native merge options determine which content is imported.</para>
/// <example>
///   <summary>Combine PDFs in the specified order.</summary>
///   <code>Merge-OfficeDocument -Path ./agenda.pdf, ./minutes.pdf -OutputPath ./board-pack.pdf -PassThru</code>
/// </example>
/// <example>
///   <summary>Combine Word files in pipeline order.</summary>
///   <code>Get-ChildItem ./weekly -Filter *.docx | Sort-Object Name | Merge-OfficeDocument -OutputPath ./monthly.docx</code>
/// </example>
/// <example>
///   <summary>Append files to a caller-owned proposal.</summary>
///   <code>Merge-OfficeDocument -Document $proposal -Path ./scope.docx, ./pricing.docx -OutputPath ./proposal.docx</code>
/// </example>
[Cmdlet("Merge", "OfficeDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(OfficeIMO.Word.WordDocument), typeof(OfficeIMO.Excel.ExcelDocument))]
public sealed class MergeOfficeDocumentCommand : PSCmdlet {
    private readonly List<string> _inputs = new();
    private object? _document;

    /// <summary>Files in merge order, or files appended to Document; wildcard matches retain provider order.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Document")]
    [SupportsWildcards]
    [ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Literal files in merge order, including FileInfo.FullName and provider PSPath properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    [ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>A single editable Word or Excel document to append to; ownership stays with the caller.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Document")]
    [ValidateNotNull]
    public IDisposable Document { get; set; } = null!;

    /// <summary>Destination with the same extension as every input; optional for a live document.</summary>
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 1, ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Document")]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>One source password reused for all files, or one password per resolved file in order.</summary>
    [Parameter]
    [AllowNull]
    [AllowEmptyCollection]
    public SecureString[]? Password { get; set; }

    /// <summary>Native Word, PDF or Excel merge options objects or property hashtables for the selected format.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public object[]? Options { get; set; }

    /// <summary>Refuse to replace an existing destination; requires OutputPath.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Emit the saved FileInfo, or the caller-owned live document for Document input.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Open the saved result through the shell after successful publication.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Show { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool live = ParameterSetName == "Document";
        if (live) {
            if (_document != null) throw new PSArgumentException("Provide exactly one live Document per merge invocation.");
            _document = OfficeMergeOptionBinding.Unwrap(Document)
                ?? throw new PSArgumentException("Document cannot be null.");
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string path in literal ? LiteralPath : Path)
            _inputs.AddRange(OfficePath.ResolveInput(SessionState, path, literal));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (_inputs.Count == 0) throw new PSArgumentException("Provide at least one input file.");
        string extension = System.IO.Path.GetExtension(_inputs[0]).ToLowerInvariant();
        var family = OfficeMergeDispatch.GetFamily(extension);
        foreach (string path in _inputs) {
            if (!extension.Equals(System.IO.Path.GetExtension(path), StringComparison.OrdinalIgnoreCase))
                throw new PSArgumentException("All merge inputs must have the same format and extension. Convert mixed inputs separately.");
        }
        if (Password != null && Password.Length != 1 && Password.Length != _inputs.Count)
            throw new PSArgumentException("Provide one Password for all files, or one Password per resolved input file.");
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        object options = OfficeMergeOptionBinding.Capture(engine, family, Options);
        string? associatedPath = _document == null ? null : OfficeMergeDispatch.ValidateDocument(_document, extension);
        string? output = OutputPath == null ? null : OfficePath.ResolveOutput(SessionState, OutputPath, directory: false);
        if (output == null && NoClobber.IsPresent) throw new PSArgumentException("NoClobber requires OutputPath.");
        if (output != null) {
            if (!extension.Equals(System.IO.Path.GetExtension(output), StringComparison.OrdinalIgnoreCase))
                throw new PSArgumentException("OutputPath must have the same format and extension as the merge inputs.");
            // Reject source replacement even with the default overwrite policy. A merge always has a distinct destination.
            foreach (string source in _inputs) RejectSourceDestination(source, output);
            if (!string.IsNullOrWhiteSpace(associatedPath)) RejectSourceDestination(associatedPath!, output);
            OfficePath.ValidateDestination(output, NoClobber.IsPresent);
        }
        if (!ShouldProcess(output ?? "Live " + family + " document", "Merge documents in input order")) return;
        OfficeMergeDispatch.Execute(_inputs, Password, family, extension, options, _document, output, NoClobber.IsPresent);
        if (Show.IsPresent) FileOpenService.Open(output!);
        if (PassThru.IsPresent) WriteObject(_document ?? (object)new FileInfo(output!));
    }

    private static void RejectSourceDestination(string source, string output) {
        if (OfficePath.AreEquivalent(source, output))
            throw new PSArgumentException("OutputPath must differ from every source document path.");
    }
}
