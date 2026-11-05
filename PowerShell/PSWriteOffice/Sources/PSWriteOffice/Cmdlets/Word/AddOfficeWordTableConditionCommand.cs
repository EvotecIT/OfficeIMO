using System.Linq;
using OfficeIMO;
using PSWriteOffice.Foundation;
using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Attaches conditional formatting logic to the current table.</summary>
/// <para>Evaluates each data row via <c>-FilterScript</c> (<c>$_</c> holds the original object) and optionally changes the table style or row shading.</para>
/// <example>
///   <summary>Shade rows above threshold.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordTableCondition -FilterScript { $_.Qty -gt 10 } -BackgroundColor '#fff4d6'</code>
///   <para>Applies a light highlight when the quantity column exceeds 10.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTableCondition", DefaultParameterSetName = "Default")]
[Alias("WordTableCondition") ]
public sealed class AddOfficeWordTableConditionCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Predicate executed per data row (uses <c>$_</c>).</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Default")]
    public ScriptBlock FilterScript { get; set; } = null!;

    /// <summary>Optional table style applied when the predicate matches.</summary>
    [Parameter]
    public WordTableStyle? TableStyle { get; set; }

    /// <summary>Row highlight color applied when the predicate matches.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>Header name of a single target column; omission targets the whole row.</summary>
    [Parameter]
    public string? Column { get; set; }
    /// <summary>Text color on matching cells.</summary>
    [Parameter, OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>Enable or explicitly disable bold on matching cells.</summary>
    [Parameter]
    public SwitchParameter Bold { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!TableStyle.HasValue && !BackgroundColor.HasValue && !Color.HasValue && !MyInvocation.BoundParameters.ContainsKey(nameof(Bold))) {
            ThrowTerminatingError(new ErrorRecord(
                new ArgumentException("Specify TableStyle or BackgroundColor."),
                "WordTableConditionNoAction",
                ErrorCategory.InvalidArgument,
                null));
            return;
        }

        var context = WordDslContext.Require(this);
        var table = context.CurrentTable ?? throw new InvalidOperationException("WordTableCondition must be used inside WordTable.");
        int? columnIndex = null;
        if (Column != null) {
            var headers = table.Rows[0].Cells.Select(cell => string.Concat(cell.Paragraphs.Select(paragraph => paragraph.Text))).ToArray();
            int found = Array.FindIndex(headers, name => string.Equals(name, Column, StringComparison.OrdinalIgnoreCase));
            if (found < 0) throw new PSArgumentException("Column must name an existing table header.", nameof(Column));
            columnIndex = found;
        }

        context.AddTableCondition(table, new WordTableConditionModel(FilterScript, TableStyle, BackgroundColor, Color, OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Bold)), columnIndex));
        WritePassThru(table);
    }

}
