namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight readback metadata for one structure element in a tagged PDF structure tree.
/// </summary>
public sealed class PdfStructureElementInfo {
    internal PdfStructureElementInfo(
        int objectNumber,
        string? structureType,
        int? parentObjectNumber,
        int? pageObjectNumber,
        string? language,
        string? alternateText,
        IReadOnlyList<int> childElementObjectNumbers,
        int markedContentReferenceCount,
        int objectReferenceCount) {
        ObjectNumber = objectNumber;
        StructureType = structureType;
        ParentObjectNumber = parentObjectNumber;
        PageObjectNumber = pageObjectNumber;
        Language = language;
        AlternateText = alternateText;
        ChildElementObjectNumbers = childElementObjectNumbers;
        MarkedContentReferenceCount = markedContentReferenceCount;
        ObjectReferenceCount = objectReferenceCount;
    }

    /// <summary>Structure element object number.</summary>
    public int ObjectNumber { get; }

    /// <summary>Structure type from /S, for example Document, P, H1, L, LI, Table, TR, TH, TD, Figure, Link, or Form.</summary>
    public string? StructureType { get; }

    /// <summary>Parent structure element or StructTreeRoot object number from /P.</summary>
    public int? ParentObjectNumber { get; }

    /// <summary>Page object number from /Pg, when present.</summary>
    public int? PageObjectNumber { get; }

    /// <summary>Language tag from /Lang, when present.</summary>
    public string? Language { get; }

    /// <summary>Alternate text from /Alt, when present.</summary>
    public string? AlternateText { get; }

    /// <summary>Child structure element object numbers found in /K.</summary>
    public IReadOnlyList<int> ChildElementObjectNumbers { get; }

    /// <summary>Number of marked-content references found in /K.</summary>
    public int MarkedContentReferenceCount { get; }

    /// <summary>Number of annotation/object references found in /K.</summary>
    public int ObjectReferenceCount { get; }

    /// <summary>True when the structure element contains at least one /K child structure reference.</summary>
    public bool HasChildElements => ChildElementObjectNumbers.Count > 0;
}
