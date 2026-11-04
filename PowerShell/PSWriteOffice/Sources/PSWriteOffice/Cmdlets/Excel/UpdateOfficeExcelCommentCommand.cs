using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Replaces the text or author of notes that match a filter.</summary>
/// <example>
/// <summary>Change the text of matching notes</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     Update-OfficeExcelComment -Sheet Assumptions -TextContains 'TBD' -Text 'Confirmed with the vendor on 24 October.'
/// }</code>
/// <para>Replaces the text of every note that says TBD on the Assumptions sheet.</para>
/// </example>
/// <example>
/// <summary>Hand notes over to a new owner</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Helpdesk.xlsx {
///     Update-OfficeExcelComment -MatchAuthor 'jnowak' -Author 'Service Desk' -Initials 'SD'
///     Update-OfficeExcelComment -Sheet Tickets -Range F40 {
///         ExcelText 'Escalated: ' -Bold -Color Danger
///         ExcelText 'vendor case 88412 opened.'
///     }
/// }</code>
/// <para>Moves every note written by a departing colleague to the team account and rewrites one escalation note with rich text.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeExcelComment", DefaultParameterSetName = "Text")]
public sealed class UpdateOfficeExcelCommentCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter()]
    public string? Range { get; set; }

    /// <summary>Existing comment author to match, ignoring case.</summary>
    [Parameter()]
    public string? MatchAuthor { get; set; }

    /// <summary>Existing text fragment to match, ignoring case.</summary>
    [Parameter()]
    public string? TextContains { get; set; }

    /// <summary>Explicitly selects every item when no filter is supplied.</summary>
    [Parameter()]
    public SwitchParameter All { get; set; }

    /// <summary>Replacement author name.</summary>
    [Parameter()]
    public string? Author { get; set; }

    /// <summary>Author initials, supplied together with Author.</summary>
    [Parameter()]
    public string? Initials { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Replacement plain text.</summary>
    [Parameter(ParameterSetName = "Text")]
    public string? Text { get; set; }

    /// <summary>Excel child commands executed in the current content scope.</summary>
    [Parameter(ParameterSetName = "RichText", Mandatory = true, Position = 0)]
    public ScriptBlock Content { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Range == null && MatchAuthor == null && TextContains == null && !All) throw new PSArgumentException("Provide a comment filter or -All.");
        if (Initials != null && Author == null) throw new PSArgumentException("Initials requires Author.");
        var document = ExcelCommandTargets.Document(Document);
        var filter = new ExcelCommentFilter { A1Range = Range, Author = MatchAuthor, TextContains = TextContains };
        int count = 0;
        var runs = Content == null ? null : ExcelRichTextRunService.ReadContent(Content);
        foreach (var sheet in ExcelSheetResolver.ResolveMany(document, Sheet, Document == null ? ExcelDslContext.Current?.CurrentSheet : null)) {
            if (runs != null) count += sheet.UpdateCommentsRichText(filter, runs, Author, Initials);
            else if (MyInvocation.BoundParameters.ContainsKey(nameof(Text))) count += sheet.UpdateComments(filter, Text!, Author, Initials);
            else {
                if (Author == null && Initials == null) throw new PSArgumentException("Supply Text, Content, Author or Initials to update comments.");
                foreach (var comment in sheet.FindComments(filter)) {
                    var single = new ExcelCommentFilter { A1Range = comment.CellReference };
                    count += comment.RichTextRuns.Count > 0
                        ? sheet.UpdateCommentsRichText(single, comment.RichTextRuns, Author, Initials)
                        : sheet.UpdateComments(single, comment.Text, Author, Initials);
                }
            }
        }
        if (PassThru) WriteObject(count);
    }

}
