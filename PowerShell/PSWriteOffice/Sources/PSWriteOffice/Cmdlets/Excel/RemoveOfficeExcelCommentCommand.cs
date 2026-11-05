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

/// <summary>Removes notes from a cell or all notes that match a filter.</summary>
/// <example>
/// <summary>Remove the note from a cell</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     Remove-OfficeExcelComment B4 -Sheet Assumptions
/// }</code>
/// <para>Deletes the note on B4 of the Assumptions sheet.</para>
/// </example>
/// <example>
/// <summary>Strip review notes before publishing</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     $removed = Remove-OfficeExcelComment -Author 'Review Bot' -PassThru
///     Write-Verbose "Removed $removed review notes"
/// }</code>
/// <para>Removes every note left by the review bot across all sheets and reports how many were removed.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeExcelComment", DefaultParameterSetName = "Default")]
public sealed class RemoveOfficeExcelCommentCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>A1 cell range on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Default", Position = 0)]
    public string? Range { get; set; }

    /// <summary>Comment author to match, ignoring case.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Author { get; set; }

    /// <summary>Text fragment to match, ignoring case.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? TextContains { get; set; }

    /// <summary>Explicitly selects every item when no filter is supplied.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter All { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Range == null && Author == null && TextContains == null && !All) throw new PSArgumentException("Provide a comment filter or -All.");
        var document = ExcelCommandTargets.Document(Document);
        var filter = new ExcelCommentFilter { A1Range = Range, Author = Author, TextContains = TextContains };
        int count = 0;

        foreach (var sheet in ExcelSheetResolver.ResolveMany(document, Sheet, Document == null ? ExcelDslContext.Current?.CurrentSheet : null)) {
            count += sheet.ClearComments(filter);
        }
        if (PassThru) WriteObject(count);
    }

}
