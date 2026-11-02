using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a PAGE field to the current header/footer.</summary>
/// <para>Typically used inside <c>WordFooter</c> to render “Page X of Y”.</para>
/// <example>
///   <summary>Footer page numbering.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>WordFooter { Add-OfficeWordPageNumber -IncludeTotal }</code>
///   <para>Outputs “Page # of #” in the footer.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordPageNumber", DefaultParameterSetName = "Default")]
[Alias("WordPageNumber")]
public sealed class AddOfficeWordPageNumberCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Include “of N” when true.</summary>
    [Parameter]
    public SwitchParameter IncludeTotal { get; set; }

    /// <summary>Optional number format.</summary>
    [Parameter]
    public WordFieldFormat? Format { get; set; }

    /// <summary>Separator when totals are included.</summary>
    [Parameter]
    public string Separator { get; set; } = " of ";

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = WordDslContext.Require(this);
        WordHeaderFooter? target = context.CurrentFooter ?? context.CurrentHeader as WordHeaderFooter;
        if (target == null) {
            throw new InvalidOperationException("WordPageNumber must be called within WordHeader or WordFooter.");
        }

        target.AddPageNumber(IncludeTotal.IsPresent, Format, Separator);
        WritePassThru(target);
    }
}