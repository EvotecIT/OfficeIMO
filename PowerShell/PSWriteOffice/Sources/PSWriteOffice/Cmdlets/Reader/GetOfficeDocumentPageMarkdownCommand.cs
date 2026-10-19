using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Projects Reader pages into citation-friendly Markdown.</summary>
/// <example>
///   <summary>Create one Markdown file with page markers.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeDocument -Path .\Handbook.pdf |
///     Get-OfficeDocumentPageMarkdown -AsString |
///     Set-Content -Path .\Handbook.pages.md</code>
///   <para>Uses OfficeIMO.Reader page projection and preserves the page provenance in each marker.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentPageMarkdown")]
[OutputType(typeof(OfficeDocumentPageMarkdown), typeof(string))]
public sealed class GetOfficeDocumentPageMarkdownCommand : PSCmdlet
{
    /// <summary>Normalized document returned by Get-OfficeDocument.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public OfficeDocumentReadResult InputObject { get; set; } = null!;

    /// <summary>Return one combined Markdown string instead of one page result per page.</summary>
    [Parameter]
    public SwitchParameter AsString { get; set; }

    /// <summary>Omit portable HTML page markers from the Markdown.</summary>
    [Parameter]
    public SwitchParameter NoPageMarkers { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var pages = InputObject.GetPageMarkdown(new OfficeDocumentPageMarkdownOptions
        {
            IncludePageMarkers = !NoPageMarkers.IsPresent
        });

        if (AsString.IsPresent)
        {
            WriteObject(string.Join(
                Environment.NewLine + Environment.NewLine,
                pages.Select(static page => page.Markdown)));
            return;
        }

        WriteObject(pages, enumerateCollection: true);
    }
}
