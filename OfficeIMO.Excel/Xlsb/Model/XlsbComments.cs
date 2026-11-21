namespace OfficeIMO.Excel.Xlsb.Model;

/// <summary>
/// A worksheet's legacy comments decoded from its Comments part (MS-XLSB 2.1.7.8) and the VML drawing part that its
/// BrtLegacyDrawing record names.
/// </summary>
internal sealed class XlsbComments {
    internal XlsbComments(string partName, string relationshipId, IReadOnlyList<string> authors, IReadOnlyList<XlsbComment> items,
        byte[] trailingRecords) {
        PartName = partName;
        RelationshipId = relationshipId;
        Authors = authors;
        Items = items;
        TrailingRecords = trailingRecords;
    }

    /// <summary>The comments part path without a leading slash.</summary>
    internal string PartName { get; }

    /// <summary>The worksheet relationship id of the comments part.</summary>
    internal string RelationshipId { get; }

    internal IReadOnlyList<string> Authors { get; }

    internal IReadOnlyList<XlsbComment> Items { get; }

    /// <summary>The FRT records before BrtEndComments, kept as encoded records.</summary>
    internal byte[] TrailingRecords { get; }

    /// <summary>The VML drawing part path without a leading slash, or null when the worksheet has no BrtLegacyDrawing.</summary>
    internal string? VmlPartName { get; set; }

    /// <summary>The worksheet relationship id that BrtLegacyDrawing stores.</summary>
    internal string? VmlRelationshipId { get; set; }
}

/// <summary>One COMMENT: the cell, the author index, the RichStr text and the native records that encoded it.</summary>
internal sealed class XlsbComment {
    internal XlsbComment(int row, int column, int authorIndex, string text, XlsbTextRun[]? runs, byte[] records) {
        Row = row;
        Column = column;
        AuthorIndex = authorIndex;
        Text = text;
        Runs = runs;
        Records = records;
    }

    /// <summary>1-based row.</summary>
    internal int Row { get; }

    /// <summary>1-based column.</summary>
    internal int Column { get; }

    internal int AuthorIndex { get; }

    internal string Text { get; }

    /// <summary>The StrRun formatting runs, or null when the text has none.</summary>
    internal XlsbTextRun[]? Runs { get; }

    /// <summary>The encoded [ACUID] BrtBeginComment [BrtCommentText] BrtEndComment records, reused while the comment is unchanged.</summary>
    internal byte[] Records { get; }
}
