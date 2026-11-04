using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Built-in heading and cover-line styles accepted by WordHeading.</summary>
public enum WordHeadingLevel {
    /// <summary>First-level heading.</summary>
    Heading1 = 1,
    /// <summary>Second-level heading.</summary>
    Heading2,
    /// <summary>Third-level heading.</summary>
    Heading3,
    /// <summary>Fourth-level heading.</summary>
    Heading4,
    /// <summary>Fifth-level heading.</summary>
    Heading5,
    /// <summary>Sixth-level heading.</summary>
    Heading6,
    /// <summary>Seventh-level heading.</summary>
    Heading7,
    /// <summary>Eighth-level heading.</summary>
    Heading8,
    /// <summary>Ninth-level heading.</summary>
    Heading9,
    /// <summary>Document title, outside the numbered heading hierarchy.</summary>
    Title,
    /// <summary>Document subtitle, outside the numbered heading hierarchy.</summary>
    Subtitle
}

/// <summary>Adds a heading using the document's built-in heading, title or subtitle style.</summary>
/// <para>Levels 1–9 participate in document navigation and tables of contents. A content block accepts WordText runs.</para>
/// <example><summary>Add a heading hierarchy.</summary>
/// <code>WordHeading 'Service review'
/// WordHeading 'Availability' -Level 2</code></example>
/// <example><summary>Add a rich heading to a live document.</summary>
/// <code>Add-OfficeWordHeading -Document $document -Level 2 { WordText 'Open incidents '; WordText '(3)' -Bold }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficeWordHeading", DefaultParameterSetName = "Text")]
[Alias("WordHeading")]
[OutputType(typeof(WordParagraph))]
public sealed class AddOfficeWordHeadingCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Heading text.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Text")]
    public string Text { get; set; } = string.Empty;
    /// <summary>Rich heading content composed with WordText.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Content")]
    public ScriptBlock? Content { get; set; }
    /// <summary>Heading level 1–9, Title or Subtitle. Defaults to Heading1.</summary>
    [Parameter]
    public WordHeadingLevel Level { get; set; } = WordHeadingLevel.Heading1;
    /// <summary>Optional horizontal alignment.</summary>
    [Parameter]
    public OfficeAlignment? Alignment { get; set; }
    /// <summary>Live document; omitted inside a Word content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!Enum.IsDefined(typeof(WordHeadingLevel), Level))
            throw new PSArgumentException("Level must be 1–9, Title or Subtitle.", nameof(Level));
        if (Alignment.HasValue && !Enum.IsDefined(typeof(OfficeAlignment), Alignment.Value))
            throw new PSArgumentException("Unsupported heading alignment.", nameof(Alignment));
        var current = WordDslContext.Current;
        if (Content != null && Document != null && current != null && !ReferenceEquals(current.Document, Document))
            throw new PSInvalidOperationException("The active Word DSL context belongs to a different document.");
        var paragraph = Document != null
            ? Document.AddParagraph(Text)
            : WordDslContext.Require(this).RequireParagraphHost().AddParagraph(Text);
        if (Level is WordHeadingLevel.Title or WordHeadingLevel.Subtitle)
            paragraph.SetStyleId(Level.ToString());
        else
            paragraph.Style = WordHeadingStyleMapper.GetHeadingStyleForLevel((int)Level);
        if (Alignment.HasValue) {
            paragraph.ParagraphAlignment = Alignment.Value switch {
                OfficeAlignment.Left or OfficeAlignment.Start => WordParagraphAlignment.Left,
                OfficeAlignment.Right or OfficeAlignment.End => WordParagraphAlignment.Right,
                OfficeAlignment.Center => WordParagraphAlignment.Center,
                OfficeAlignment.Justify => WordParagraphAlignment.Both,
                _ => throw new PSArgumentException("Unsupported heading alignment.", nameof(Alignment))
            };
        }
        if (Content != null) {
            using var ownedContext = current == null && Document != null ? WordDslContext.Enter(Document) : null;
            using (WordDslContext.Require(this).Push(paragraph)) Content.InvokeReturnAsIs();
        }
        WritePassThru(paragraph);
    }
}
