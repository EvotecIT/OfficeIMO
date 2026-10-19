using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a discoverable Word revision filter for Resolve-OfficeWordRevision.</summary>
/// <example>
///   <summary>Accept only table revisions from one author.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$filter = New-OfficeWordRevisionFilter -Author 'Alex' -InTable
/// Resolve-OfficeWordRevision -Path .\Review.docx -Action Accept -Filter $filter</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordRevisionFilter")]
[OutputType(typeof(WordRevisionFilter))]
public sealed class NewOfficeWordRevisionFilterCommand : PSCmdlet {
    /// <summary>Revision author.</summary>
    [Parameter] public string? Author { get; set; }
    /// <summary>Revision identifier.</summary>
    [Parameter] public string? RevisionId { get; set; }
    /// <summary>Revision operation type.</summary>
    [Parameter] public WordReviewRevisionType? RevisionType { get; set; }
    /// <summary>Earliest revision date.</summary>
    [Parameter] public System.DateTime? DateFrom { get; set; }
    /// <summary>Latest revision date.</summary>
    [Parameter] public System.DateTime? DateTo { get; set; }
    /// <summary>Word part or container location kind.</summary>
    [Parameter] public WordReviewLocationKind? LocationKind { get; set; }
    /// <summary>Package part URI.</summary>
    [Parameter] public string? PartUri { get; set; }
    /// <summary>Limit results to revisions inside tables.</summary>
    [Parameter] public SwitchParameter InTable { get; set; }
    /// <summary>Limit results to revisions outside tables.</summary>
    [Parameter] public SwitchParameter NotInTable { get; set; }
    /// <summary>Limit results to revisions inside content controls.</summary>
    [Parameter] public SwitchParameter InContentControl { get; set; }
    /// <summary>Limit results to revisions outside content controls.</summary>
    [Parameter] public SwitchParameter NotInContentControl { get; set; }
    /// <summary>Limit results to revisions inside text boxes.</summary>
    [Parameter] public SwitchParameter InTextBox { get; set; }
    /// <summary>Limit results to revisions outside text boxes.</summary>
    [Parameter] public SwitchParameter NotInTextBox { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        ValidatePair(nameof(InTable), InTable, nameof(NotInTable), NotInTable);
        ValidatePair(nameof(InContentControl), InContentControl, nameof(NotInContentControl), NotInContentControl);
        ValidatePair(nameof(InTextBox), InTextBox, nameof(NotInTextBox), NotInTextBox);
        var filter = new WordRevisionFilter {
            Author = Author,
            RevisionId = RevisionId,
            RevisionType = RevisionType,
            DateFrom = DateFrom,
            DateTo = DateTo,
            LocationKind = LocationKind,
            PartUri = PartUri
        };
        if (InTable.IsPresent || NotInTable.IsPresent) filter.IsInTable = InTable.IsPresent;
        if (InContentControl.IsPresent || NotInContentControl.IsPresent) filter.IsInContentControl = InContentControl.IsPresent;
        if (InTextBox.IsPresent || NotInTextBox.IsPresent) filter.IsInTextBox = InTextBox.IsPresent;
        WriteObject(filter);
    }

    private static void ValidatePair(string includeName, SwitchParameter include, string excludeName, SwitchParameter exclude) {
        if (include.IsPresent && exclude.IsPresent) throw new PSArgumentException($"-{includeName} and -{excludeName} cannot be used together.");
    }
}
