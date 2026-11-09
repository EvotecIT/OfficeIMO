using System.Management.Automation;
using OfficeIMO.Adf;
using PSWriteOffice.Models.Confluence;
using OfficeIMO.Confluence;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Reads one page or streams a filtered Confluence Cloud page listing.</summary>
/// <example>
/// <summary>Read a page as Markdown with fidelity evidence.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Get-OfficeConfluencePage -Session $session -PageId 12345 -AsMarkdown</code>
/// <para>Returns the converted Markdown value and its ADF conversion report.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeConfluencePage", DefaultParameterSetName = ParameterSetList)]
[OutputType(typeof(ConfluencePage), typeof(ConfluencePageBatch), typeof(ConfluenceContentConversionResult<string>))]
public sealed class GetOfficeConfluencePageCommand : AsyncPSCmdlet
{
    private const string ParameterSetById = "ById";
    private const string ParameterSetList = "List";

    /// <summary>Configured Confluence session.</summary>
    [Parameter(Mandatory = true)]
    public ConfluenceSession Session { get; set; } = null!;

    /// <summary>Page identifier.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetById)]
    [ValidateNotNullOrEmpty]
    public string PageId { get; set; } = string.Empty;

    /// <summary>Optional space identifier used when listing pages.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    public string? SpaceId { get; set; }

    /// <summary>Optional exact title used when listing pages.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    public string? Title { get; set; }

    /// <summary>Optional cursor at which to resume a page listing.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    public string? Cursor { get; set; }

    /// <summary>Maximum pages requested in each listing batch.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    [ValidateRange(1, 250)]
    public int Limit { get; set; } = 25;

    /// <summary>Return listing batches rather than enumerating their pages.</summary>
    [Parameter(ParameterSetName = ParameterSetList)]
    public SwitchParameter AsBatch { get; set; }

    /// <summary>Body representation requested from Confluence.</summary>
    [Parameter]
    public ConfluenceBodyFormat BodyFormat { get; set; } = ConfluenceBodyFormat.AtlasDocFormat;

    /// <summary>Native page, Markdown or HTML projection with fidelity evidence.</summary>
    [Parameter]
    public OfficeConfluencePageOutput As { get; set; } = OfficeConfluencePageOutput.Page;

    /// <summary>Throw when a requested projection reports reduced fidelity.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        if (AsBatch.IsPresent && As != OfficeConfluencePageOutput.Page)
            throw new PSArgumentException("-AsBatch returns native batches; per-page projections require -AsBatch:$false.");

        using var client = Session.CreateClient();
        if (ParameterSetName == ParameterSetById)
        {
            WritePage(await client.GetPageAsync(PageId, BodyFormat, CancelToken).ConfigureAwait(false));
            return;
        }

        var cursor = Cursor;
        var observed = new HashSet<string>(StringComparer.Ordinal);
        if (cursor != null)
        {
            observed.Add(cursor);
        }

        do
        {
            var requestedCursor = cursor;
            var batch = await client.GetPagesAsync(new ConfluencePageQuery
            {
                SpaceId = SpaceId,
                Title = Title,
                Cursor = requestedCursor,
                Limit = Limit,
                BodyFormat = BodyFormat
            }, CancelToken).ConfigureAwait(false);

            if (AsBatch.IsPresent)
            {
                WriteObject(batch);
            }
            else
            {
                foreach (var page in batch.Pages)
                {
                    WritePage(page);
                }
            }

            cursor = batch.NextCursor;
            if (cursor != null &&
                (string.Equals(cursor, requestedCursor, StringComparison.Ordinal) || !observed.Add(cursor)))
            {
                throw new InvalidOperationException("Confluence returned a repeated page cursor and cannot make progress.");
            }
        }
        while (cursor != null);
    }

    private void WritePage(ConfluencePage page)
    {
        if (As == OfficeConfluencePageOutput.Markdown)
        {
            var result = ConfluenceContentConverter.ToMarkdown(page);
            EnsureNoLoss(result.Report);
            WriteObject(result);
        }
        else if (As == OfficeConfluencePageOutput.Html)
        {
            var result = ConfluenceContentConverter.ToHtml(page);
            EnsureNoLoss(result.Report);
            WriteObject(result);
        }
        else
        {
            WriteObject(page);
        }
    }

    private void EnsureNoLoss(AdfConversionReport report)
    {
        if (FailOnLoss.IsPresent && !report.IsLossless)
        {
            throw new InvalidOperationException("Confluence content projection reported one or more fidelity losses.");
        }
    }
}
