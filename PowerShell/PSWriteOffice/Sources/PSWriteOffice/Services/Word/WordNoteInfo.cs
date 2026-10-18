using System;
using System.Collections.Generic;

namespace PSWriteOffice.Services.Word;

/// <summary>Describes a Word footnote or endnote in a document-safe snapshot.</summary>
public sealed class WordNoteInfo
{
    /// <summary>Initializes a note snapshot.</summary>
    public WordNoteInfo(string noteType, long? referenceId, string? parentText, IReadOnlyList<string> paragraphs)
    {
        NoteType = noteType;
        ReferenceId = referenceId;
        ParentText = parentText;
        Paragraphs = paragraphs ?? Array.Empty<string>();
        Text = string.Join(Environment.NewLine, Paragraphs);
    }

    /// <summary>Footnote or endnote.</summary>
    public string NoteType { get; }

    /// <summary>Open XML note reference identifier.</summary>
    public long? ReferenceId { get; }

    /// <summary>Text of the paragraph that owns the note reference.</summary>
    public string? ParentText { get; }

    /// <summary>Text paragraphs inside the note.</summary>
    public IReadOnlyList<string> Paragraphs { get; }

    /// <summary>Combined note text.</summary>
    public string Text { get; }
}
