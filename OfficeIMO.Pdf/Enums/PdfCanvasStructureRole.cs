namespace OfficeIMO.Pdf;

/// <summary>Semantic container roles available for absolute canvas content.</summary>
public enum PdfCanvasStructureRole {
    /// <summary>Document section.</summary>
    Section,
    /// <summary>Generic semantic division.</summary>
    Division,
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
