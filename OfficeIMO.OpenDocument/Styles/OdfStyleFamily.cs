namespace OfficeIMO.OpenDocument;

/// <summary>Style families shared by OpenDocument text, spreadsheet, and presentation documents.</summary>
public enum OdfStyleFamily {
    /// <summary>Inline text formatting.</summary>
    Text,
    /// <summary>Paragraph formatting.</summary>
    Paragraph,
    /// <summary>Table formatting.</summary>
    Table,
    /// <summary>Table row formatting.</summary>
    TableRow,
    /// <summary>Table column formatting.</summary>
    TableColumn,
    /// <summary>Table cell formatting.</summary>
    TableCell,
    /// <summary>Drawing and presentation graphic formatting.</summary>
    Graphic,
    /// <summary>Presentation page formatting.</summary>
    Presentation,
    /// <summary>Chart formatting.</summary>
    Chart
}
