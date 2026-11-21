using System.Management.Automation;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models;

namespace PSWriteOffice.Cmdlets.Security;

/// <summary>Writes a copy of a file without its metadata and hidden data, and reports what happened to each item.</summary>
/// <para>The input is never changed. The cleaned copy is written as &lt;name&gt;.cleaned&lt;ext&gt; beside the input unless
/// OutputPath or OutputDirectory says otherwise. An existing output file is replaced, as with Convert-OfficeDocument and
/// Optimize-OfficePdf; NoClobber refuses it instead.</para>
/// <para>Without Options the Default categories are removed: document and custom properties, comments, personal identities,
/// speaker notes, the thumbnail, custom XML, image metadata, printer settings, external references, document variables,
/// revision identifiers and active content such as scripts. Tracked changes, hidden content, provenance, attachments and
/// bookmarks are kept until New-OfficeHiddenDataOptions selects them. Every finding reports an Outcome; a Retained finding
/// carries a Reason and a Note that say why it stayed.</para>
/// <para>Saving refuses to invalidate an Office package signature unless RemoveInvalidatedSignatures is present. Signed PDFs
/// are always refused.</para>
/// <example>
///   <summary>Write Proposal.cleaned.docx with the default categories removed.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Remove-OfficeHiddenData -Path .\Proposal.docx</code>
/// </example>
/// <example>
///   <summary>Clean a folder into a share folder, also removing Content Credentials and AI labels.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeHiddenDataOptions -Kind Default, Provenance
/// Get-ChildItem .\Outgoing -File | Remove-OfficeHiddenData -OutputDirectory .\Share -Options $options</code>
/// </example>
/// <example>
///   <summary>Accept tracked changes, and keep an earlier cleaned copy if one exists.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeHiddenDataOptions -TrackedChanges AcceptAll
/// Remove-OfficeHiddenData -Path .\Contract.docx -OutputPath .\Contract-final.docx -Options $options -NoClobber</code>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeHiddenData", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeHiddenDataFileReport))]
public sealed class RemoveOfficeHiddenDataCommand : OfficeCmdlet {
    private readonly List<string> _inputs = new();

    /// <summary>Files to clean; supports wildcards.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "PathDirectory")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();

    /// <summary>Exact file paths, including FileInfo.FullName and provider PSPath pipeline properties.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPathDirectory")]
    [Alias("PSPath", "FullName"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();

    /// <summary>The cleaned copy for a single input. It must keep the input's extension.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateNotNullOrEmpty]
    public string? OutputPath { get; set; }

    /// <summary>Folder that receives each cleaned copy under its original file name; created when missing.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "PathDirectory")]
    [Parameter(Mandatory = true, ParameterSetName = "LiteralPathDirectory")]
    [ValidateNotNullOrEmpty]
    public string? OutputDirectory { get; set; }

    /// <summary>Categories to remove and how to treat tracked changes and provenance, from New-OfficeHiddenDataOptions. Defaults to the Default categories.</summary>
    [Parameter]
    [OfficeOptionsTransformation(typeof(OfficeHiddenDataOptions))]
    public OfficeHiddenDataOptions? Options { get; set; }

    /// <summary>Allow cleaning to remove Office package signatures that it invalidates; otherwise a signed package is refused.</summary>
    [Parameter]
    public SwitchParameter RemoveInvalidatedSignatures { get; set; }

    /// <summary>Refuse to replace an existing output file.</summary>
    [Parameter]
    public SwitchParameter NoClobber { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        bool literal = ParameterSetName.StartsWith("LiteralPath", StringComparison.Ordinal);
        _inputs.AddRange(OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (OutputPath != null && _inputs.Count > 1) {
            InvalidArgument(new PSArgumentException("OutputPath accepts one input file. Use OutputDirectory for several files."), OutputPath);
        }
        string? explicitOutput = OutputPath == null ? null : ResolveOutputPath(OutputPath);
        string? directory = OutputDirectory == null ? null : ResolveOutputDirectory(OutputDirectory);
        var policy = RemoveInvalidatedSignatures ? OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures : OfficeSignatureMutationPolicy.BlockSave;
        var written = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string input in _inputs) {
            if (Stopping) return;
            if (!OfficeDocument.SupportsHiddenData(input)) {
                ItemError(OfficeHiddenDataFiles.Unsupported(input), "hidden-data-unsupported-format", input);
                continue;
            }
            string output = explicitOutput ?? (directory == null
                ? System.IO.Path.Combine(System.IO.Path.GetDirectoryName(input)!, System.IO.Path.GetFileNameWithoutExtension(input) + ".cleaned" + System.IO.Path.GetExtension(input))
                : System.IO.Path.Combine(directory, System.IO.Path.GetFileName(input)));
            try {
                if (OfficePath.AreEquivalent(input, output)) throw new PSArgumentException("The cleaned copy must not replace its input: " + input);
                if (!string.Equals(System.IO.Path.GetExtension(input), System.IO.Path.GetExtension(output), StringComparison.OrdinalIgnoreCase))
                    throw new PSArgumentException("The cleaned copy must keep the input's extension: " + output);
                if (!written.Add(output)) throw new IOException("Another input in this run already wrote: " + output);
                OfficePath.ValidateDestination(output, refuseExisting: NoClobber);
            } catch (Exception error) when (error is ArgumentException or IOException) {
                ItemError(error, "hidden-data-output-rejected", input);
                continue;
            }
            if (!ShouldProcess(output, "Write a copy of '" + input + "' without hidden data")) continue;
            try {
                OfficeHiddenDataResult result = OfficeDocument.RemoveHiddenData(File.ReadAllBytes(input), input, Options, policy);
                string? parent = System.IO.Path.GetDirectoryName(output);
                if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
                byte[] cleaned = result.ToArray();
                OfficeOutput.WriteNew(output, stream => stream.Write(cleaned, 0, cleaned.Length), !NoClobber);
                WriteObject(new OfficeHiddenDataFileReport(result.Report, input, output));
            } catch (Exception error) when (error is not (ActionPreferenceStopException or PipelineStoppedException)) {
                ItemError(error, "hidden-data-removal-failed", input);
            }
        }
    }
}
