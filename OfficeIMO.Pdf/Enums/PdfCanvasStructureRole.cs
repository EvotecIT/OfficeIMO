namespace OfficeIMO.Pdf;

/// <summary>Semantic container roles available for absolute canvas content.</summary>
public enum PdfCanvasStructureRole {
    /// <summary>Document section.</summary>
    Section,
    /// <summary>Generic semantic division.</summary>
    Division,
    /// <summary>Paragraph containing one or more positioned text fragments.</summary>
    Paragraph,
    /// <summary>Level-one heading containing one or more positioned text fragments.</summary>
    Heading1,
    /// <summary>Level-two heading containing one or more positioned text fragments.</summary>
    Heading2,
    /// <summary>Level-three heading containing one or more positioned text fragments.</summary>
    Heading3,
    /// <summary>Level-four heading containing one or more positioned text fragments.</summary>
    Heading4,
    /// <summary>Level-five heading containing one or more positioned text fragments.</summary>
    Heading5,
    /// <summary>Level-six heading containing one or more positioned text fragments.</summary>
    Heading6,
    /// <summary>Ordered or unordered list container.</summary>
    List,
    /// <summary>One list item.</summary>
    ListItem,
    /// <summary>List item label or marker.</summary>
    ListLabel,
    /// <summary>List item body.</summary>
    ListBody,
    /// <summary>Table container.</summary>
    Table,
    /// <summary>Table row.</summary>
    TableRow,
    /// <summary>Table header cell.</summary>
    TableHeaderCell,
    /// <summary>Table data cell.</summary>
    TableCell,
    /// <summary>Caption associated with a figure or table.</summary>
    Caption
}

/// <summary>Scope of a tagged table header cell.</summary>
public enum PdfCanvasTableHeaderScope {
    /// <summary>Header applies to its row.</summary>
    Row,
    /// <summary>Header applies to its column.</summary>
    Column,
    /// <summary>Header applies to both its row and column.</summary>
    Both
}
