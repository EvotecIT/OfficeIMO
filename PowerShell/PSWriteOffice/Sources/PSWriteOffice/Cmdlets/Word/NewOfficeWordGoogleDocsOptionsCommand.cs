using System.Management.Automation;
using OfficeIMO.Word.GoogleDocs;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates native Word-to-Google-Docs planning and export settings without contacting Google.</summary>
/// <example><summary>Prepare a guarded replacement in a selected document tab.</summary>
/// <code>$options = New-OfficeWordGoogleDocsOptions -ExistingFileId $fileId -TabStrategy SelectedTab -TabId $tabId -ExpectedRevisionId $revisionId
/// Export-OfficeWordGoogleDocument -Document $document -Session $session -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeWordGoogleDocsOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(GoogleDocsSaveOptions))]
public sealed class NewOfficeWordGoogleDocsOptionsCommand : OfficeCmdlet {
    /// <summary>Title of a newly created document; replacement retains the existing title.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateNotNullOrEmpty]
    public string? Title { get; set; }
    /// <summary>Destination Drive folder identifier.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? FolderId { get; set; }
    /// <summary>Shared-drive identifier; omission uses the native session location policy.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? DriveId { get; set; }
    /// <summary>Existing document to replace instead of creating a new document.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ExistingFileId { get; set; }
    /// <summary>Inline image policy. Temporary public staging requires the native session's explicit authorization.</summary>
    [Parameter]
    public GoogleDocsInlineImageMode InlineImageMode { get; set; }
    /// <summary>How Word comments are transferred; omission preserves the native default.</summary>
    [Parameter]
    public GoogleDocsCommentMode Comments { get; set; }
    /// <summary>First, selected or every tab when reading or replacing content.</summary>
    [Parameter]
    public GoogleDocsTabStrategy TabStrategy { get; set; }
    /// <summary>Tab identifier for SelectedTab.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? TabId { get; set; }
    /// <summary>Revision conflict handling; omission preserves exact revision checks.</summary>
    [Parameter]
    public GoogleDocsRevisionConflictMode ConflictMode { get; set; }
    /// <summary>Revision captured by an earlier read, checked by the native exporter.</summary>
    [Parameter, ValidateNotNullOrEmpty]
    public string? ExpectedRevisionId { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new GoogleDocsSaveOptions();
        if (Title != null) options.Title = Title;
        if (FolderId != null) options.Location.FolderId = FolderId;
        if (DriveId != null) options.Location.DriveId = DriveId;
        if (ExistingFileId != null) options.Location.ExistingFileId = ExistingFileId;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(InlineImageMode))) options.InlineImageMode = InlineImageMode;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Comments))) options.Comments = Comments;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(TabStrategy))) options.Tabs.Strategy = TabStrategy;
        if (TabId != null) options.Tabs.TabId = TabId;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(ConflictMode))) options.Replace.ConflictMode = ConflictMode;
        if (ExpectedRevisionId != null) options.Replace.ExpectedRevisionId = ExpectedRevisionId;
        WriteObject(options);
    }
}
