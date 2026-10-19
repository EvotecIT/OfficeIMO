using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds content to a section header.</summary>
/// <para>Ensures the requested header part (default/first/even) exists and executes the DSL scriptblock within it.</para>
/// <example>
///   <summary>Create a default header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordSection { Add-OfficeWordHeader { Add-OfficeWordParagraph -Text 'Confidential' -Style Heading3 } }</code>
///   <para>Creates a section header that prints “Confidential”.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordHeader")]
[Alias("WordHeader")]
public sealed class AddOfficeWordHeaderCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>The header type to modify.</summary>
    [Parameter]
    public WordHeaderFooterType Type { get; set; } = WordHeaderFooterType.Default;

    /// <summary>DSL scriptblock to execute inside the header.</summary>
    [Parameter(Position = 0)]
    public ScriptBlock? Content { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = WordDslContext.Require(this);
        var section = context.RequireSection();
        var header = section.GetOrCreateHeader(Type);

        using (context.Push(header)) {
            Content?.InvokeReturnAsIs();
        }
        WritePassThru(header);
    }
}