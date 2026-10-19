using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates a Markdown document using a DSL scriptblock.</summary>
/// <para>Runs the scriptblock against a Markdown document and saves it to disk unless <c>-NoSave</c> is specified.</para>
/// <example>
///   <summary>Create a Markdown document with headings and a table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\README.md { MarkdownHeading -Level 1 -Text 'Report'; MarkdownTable -InputObject $data }</code>
///   <para>Creates a README file with a heading and table content.</para>
/// </example>
/// <example>
///   <summary>Create a report with multiple tables.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Report.md {
///     MarkdownHeading -Level 1 -Text 'Summary'
///     MarkdownTable -InputObject $summary
///     MarkdownHeading -Level 2 -Text 'Details'
///     MarkdownTable -InputObject $details
///   }</code>
///   <para>Creates a report with two tables separated by headings.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdown", SupportsShouldProcess = true)]
[Alias("OfficeMarkdown", "MarkdownNew")]
[OutputType(typeof(FileInfo), typeof(MarkdownDoc))]
public sealed class NewOfficeMarkdownCommand : PSCmdlet
    , IMarkdownWriteOptionSource {
    /// <summary>Destination path for the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "OutputPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>DSL scriptblock describing Markdown content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Optional Markdown writer options.</summary>
    [Parameter]
    public MarkdownWriteOptions? WriteOptions { get; set; }

    /// <summary>Friendly Markdown writer profile.</summary>
    [Parameter]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }

    /// <summary>Controls how Markdown images are serialized.</summary>
    [Parameter]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }

    /// <summary>Markdown line ending: CRLF, LF, CR, or a literal line ending string.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeLineEndingArgumentCompleter))]
    public string? LineEnding { get; set; }

    /// <summary>Unordered list marker: '-', '*', or '+'.</summary>
    [Parameter]
    public string? UnorderedListMarker { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var fullPath = GetResolvedPath();
        if (!NoSave.IsPresent && !PdfCommandUtilities.ShouldWrite(this, fullPath, "Write new Markdown document")) {
            return;
        }

        var document = MarkdownDoc.Create();
        if (Content != null) {
            using (MarkdownDslContext.Enter(document)) {
                Content.InvokeReturnAsIs();
            }
        }

        if (NoSave.IsPresent) {
            WriteObject(document);
            return;
        }

        var directory = System.IO.Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(fullPath, document.ToMarkdown(MarkdownOptionUtilities.BuildWriteOptions(this)), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(fullPath));
        }
    }

    private string GetResolvedPath() {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        return System.IO.Path.IsPathRooted(providerPath)
            ? providerPath
            : System.IO.Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

}
