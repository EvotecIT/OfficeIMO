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
[Cmdlet(VerbsCommon.Add, "OfficeWordTableCondition")]
[Alias("WordTableCondition")]
public sealed class AddOfficeWordTableConditionCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Predicate executed per data row (uses <c>$_</c>).</summary>
    [Parameter(Mandatory = true)]
    public ScriptBlock FilterScript { get; set; } = null!;

    /// <summary>Optional table style applied when the predicate matches.</summary>
    [Parameter]
    public WordTableStyle? TableStyle { get; set; }

    /// <summary>Row highlight color applied when the predicate matches.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!TableStyle.HasValue && string.IsNullOrWhiteSpace(BackgroundColor)) {
            ThrowTerminatingError(new ErrorRecord(
                new ArgumentException("Specify TableStyle or BackgroundColor."),
                "WordTableConditionNoAction",
                ErrorCategory.InvalidArgument,
                null));
            return;
        }

        var context = WordDslContext.Require(this);
        var table = context.CurrentTable ?? throw new InvalidOperationException("WordTableCondition must be used inside WordTable.");
        var normalizedColor = NormalizeColor(BackgroundColor);

        context.AddTableCondition(table, new WordTableConditionModel(FilterScript, TableStyle, normalizedColor));
        WritePassThru(table);
    }

    private static string? NormalizeColor(string? color) {
        if (string.IsNullOrWhiteSpace(color)) {
            return null;
        }

        return OfficeColor.Parse(color!).ToRgbHex().ToLowerInvariant();
    }
}
