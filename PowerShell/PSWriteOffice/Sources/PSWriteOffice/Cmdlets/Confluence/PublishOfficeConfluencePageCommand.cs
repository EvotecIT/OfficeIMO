using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Adf;
using OfficeIMO.Confluence;

namespace PSWriteOffice.Cmdlets.Confluence;

/// <summary>Plans, creates, or updates a Confluence Cloud page.</summary>
/// <example>
/// <summary>Inspect the exact create request without contacting Confluence.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Publish-OfficeConfluencePage -SpaceId 42 -Title 'Daily status' -Content $markdown -PlanOnly</code>
/// <para>Returns a serializable request plan and performs no network operation.</para>
/// </example>
/// <example>
/// <summary>Update a page using optimistic versioning.</summary>
/// <prefix>PS&gt; </prefix>
/// <code>Publish-OfficeConfluencePage -Session $session -PageId 123 -Title 'Daily status' -VersionNumber 8 -Content $markdown -VersionMessage 'automation refresh'</code>
/// <para>The version number must be the next Confluence page version.</para>
/// </example>
[Cmdlet(VerbsData.Publish, "OfficeConfluencePage", DefaultParameterSetName = ParameterSetCreate, SupportsShouldProcess = true)]
[OutputType(typeof(ConfluencePageWritePlan), typeof(ConfluencePage))]
public sealed class PublishOfficeConfluencePageCommand : AsyncPSCmdlet
{
    private const string ParameterSetCreate = "Create";
    private const string ParameterSetUpdate = "Update";
    private readonly List<string> _contentRecords = new();

    /// <summary>Configured session required for live create or update operations.</summary>
    [Parameter]
    public ConfluenceSession? Session { get; set; }

    /// <summary>Space identifier for a new page.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCreate)]
    [ValidateNotNullOrEmpty]
    public string SpaceId { get; set; } = string.Empty;

    /// <summary>Optional parent page identifier for a new page.</summary>
    [Parameter(ParameterSetName = ParameterSetCreate)]
    public string? ParentId { get; set; }

    /// <summary>Page identifier for an update.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetUpdate)]
    [ValidateNotNullOrEmpty]
    public string PageId { get; set; } = string.Empty;

    /// <summary>Next positive page version number for an update.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetUpdate)]
    [ValidateRange(1, int.MaxValue)]
    public int VersionNumber { get; set; }

    /// <summary>Optional version message for an update.</summary>
    [Parameter(ParameterSetName = ParameterSetUpdate)]
    public string? VersionMessage { get; set; }

    /// <summary>Page title.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string Title { get; set; } = string.Empty;

    /// <summary>Page content in the representation selected by ContentFormat.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [AllowEmptyString]
    public string Content { get; set; } = string.Empty;

    /// <summary>Representation of Content.</summary>
    [Parameter]
    public OfficeConfluenceContentFormat ContentFormat { get; set; } = OfficeConfluenceContentFormat.Markdown;

    /// <summary>Confluence representation to publish.</summary>
    [Parameter]
    public ConfluenceBodyFormat BodyFormat { get; set; } = ConfluenceBodyFormat.AtlasDocFormat;

    /// <summary>Return the exact write plan without contacting Confluence.</summary>
    [Parameter]
    public SwitchParameter PlanOnly { get; set; }

    /// <summary>Throw when conversion reports reduced fidelity.</summary>
    [Parameter]
    public SwitchParameter FailOnLoss { get; set; }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync()
    {
        _contentRecords.Add(Content);
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync()
    {
        string content = string.Join(Environment.NewLine, _contentRecords);
        var converted = ConvertContent(content);
        if (FailOnLoss.IsPresent && !converted.Report.IsLossless)
        {
            throw new InvalidOperationException("Confluence page conversion reported one or more fidelity losses.");
        }

        if (ParameterSetName == ParameterSetCreate)
        {
            var request = new ConfluencePageCreateRequest
            {
                SpaceId = SpaceId,
                ParentId = ParentId,
                Title = Title,
                Body = converted.Body
            };
            if (PlanOnly.IsPresent)
            {
                WriteObject(ConfluenceClient.PlanCreatePage(request));
                return;
            }

            EnsureLiveSession();
            if (!ShouldProcess(Title, "Create Confluence page"))
            {
                return;
            }

            using var client = Session!.CreateClient();
            WriteObject(await client.CreatePageAsync(request, CancelToken).ConfigureAwait(false));
            return;
        }

        var update = new ConfluencePageUpdateRequest
        {
            PageId = PageId,
            Title = Title,
            VersionNumber = VersionNumber,
            VersionMessage = VersionMessage,
            Body = converted.Body
        };
        if (PlanOnly.IsPresent)
        {
            WriteObject(ConfluenceClient.PlanUpdatePage(update));
            return;
        }

        EnsureLiveSession();
        if (!ShouldProcess(PageId, $"Update Confluence page to version {VersionNumber}"))
        {
            return;
        }

        using var updateClient = Session!.CreateClient();
        WriteObject(await updateClient.UpdatePageAsync(update, CancelToken).ConfigureAwait(false));
    }

    private (ConfluencePageBody Body, AdfConversionReport Report) ConvertContent(string content)
    {
        if (ContentFormat == OfficeConfluenceContentFormat.Markdown)
        {
            var result = ConfluenceContentConverter.FromMarkdown(content, BodyFormat);
            return (result.Value, result.Report);
        }

        if (ContentFormat is OfficeConfluenceContentFormat.Html or OfficeConfluenceContentFormat.Storage)
        {
            var result = ConfluenceContentConverter.FromHtml(content, BodyFormat);
            return (result.Value, result.Report);
        }

        var document = AdfDocument.Parse(content);
        var validation = document.Validate();
        if (!validation.IsValid)
        {
            throw new ArgumentException("Atlas Document Format content is structurally invalid.", nameof(Content));
        }

        if (BodyFormat == ConfluenceBodyFormat.AtlasDocFormat)
        {
            return (new ConfluencePageBody { Representation = "atlas_doc_format", Value = document.ToJson() }, new AdfConversionReport(Array.Empty<AdfConversionDiagnostic>()));
        }

        var html = AdfConverter.ToHtml(document);
        return (new ConfluencePageBody { Representation = "storage", Value = html.Value }, html.Report);
    }

    private void EnsureLiveSession()
    {
        if (Session == null)
        {
            throw new PSInvalidOperationException("Provide -Session for a live operation, or use -PlanOnly.");
        }
    }
}
