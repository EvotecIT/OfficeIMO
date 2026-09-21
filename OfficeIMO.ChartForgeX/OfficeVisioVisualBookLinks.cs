using System;
using System.Collections.Generic;
using OfficeIMO.Visio;

namespace OfficeIMO.ChartForgeX;

/// <summary>
/// Describes navigation between two native entities on different pages in a Visio book.
/// Page numbers are one-based so they can be used directly with report page indexes.
/// </summary>
public sealed class OfficeVisioVisualBookLink {
    /// <summary>Creates a cross-page link between two projected entity identifiers.</summary>
    public OfficeVisioVisualBookLink(
        int sourcePageNumber,
        string sourceEntityId,
        int targetPageNumber,
        string targetEntityId,
        string? relationshipId = null) {
        SourcePageNumber = sourcePageNumber;
        SourceEntityId = sourceEntityId;
        TargetPageNumber = targetPageNumber;
        TargetEntityId = targetEntityId;
        RelationshipId = relationshipId;
    }

    /// <summary>Gets the one-based page containing the source entity.</summary>
    public int SourcePageNumber { get; }

    /// <summary>Gets the projected source shape identifier.</summary>
    public string SourceEntityId { get; }

    /// <summary>Gets the one-based page containing the target entity.</summary>
    public int TargetPageNumber { get; }

    /// <summary>Gets the projected target shape identifier.</summary>
    public string TargetEntityId { get; }

    /// <summary>Gets the optional source relationship identifier retained with navigation metadata.</summary>
    public string? RelationshipId { get; }

    /// <summary>Gets or sets the forward-link description. A page and entity description is generated when omitted.</summary>
    public string? Description { get; set; }

    /// <summary>Gets or sets the reciprocal-link description. A page and entity description is generated when omitted.</summary>
    public string? ReturnDescription { get; set; }
}

/// <summary>Controls bounded cross-page navigation in a native Visio book.</summary>
public sealed class OfficeVisioVisualBookOptions {
    /// <summary>Gets or sets the maximum distinct page links attached to one entity. The default is 12.</summary>
    public int MaximumNavigationLinksPerEntity { get; set; } = 12;

    /// <summary>Gets or sets whether every forward link also adds a link back from the target entity. The default is true.</summary>
    public bool IncludeReturnLinks { get; set; } = true;

    /// <summary>Gets or sets whether navigation targets, relationship identifiers, and omitted counts are retained as Shape Data. The default is true.</summary>
    public bool IncludeNavigationShapeData { get; set; } = true;

    internal void Validate() {
        if (MaximumNavigationLinksPerEntity < 1) {
            throw new ArgumentOutOfRangeException(nameof(MaximumNavigationLinksPerEntity), "The per-entity navigation limit must be positive.");
        }
    }
}

/// <summary>Describes one internal page hyperlink applied to a projected Visio entity.</summary>
public sealed class OfficeVisioVisualBookNavigationResult {
    internal OfficeVisioVisualBookNavigationResult(
        int sourcePageNumber,
        string sourceEntityId,
        int targetPageNumber,
        string targetEntityId,
        IReadOnlyList<string> relationshipIds,
        bool isReturnLink,
        VisioHyperlink hyperlink) {
        SourcePageNumber = sourcePageNumber;
        SourceEntityId = sourceEntityId;
        TargetPageNumber = targetPageNumber;
        TargetEntityId = targetEntityId;
        RelationshipIds = relationshipIds;
        IsReturnLink = isReturnLink;
        Hyperlink = hyperlink;
    }

    /// <summary>Gets the one-based page containing the linked entity.</summary>
    public int SourcePageNumber { get; }

    /// <summary>Gets the linked entity identifier.</summary>
    public string SourceEntityId { get; }

    /// <summary>Gets the one-based destination page.</summary>
    public int TargetPageNumber { get; }

    /// <summary>Gets the destination entity identifier.</summary>
    public string TargetEntityId { get; }

    /// <summary>Gets source relationship identifiers represented by this deduplicated navigation link.</summary>
    public IReadOnlyList<string> RelationshipIds { get; }

    /// <summary>Gets whether this navigation was generated as a reciprocal return link.</summary>
    public bool IsReturnLink { get; }

    /// <summary>Gets the native ShapeSheet hyperlink row.</summary>
    public VisioHyperlink Hyperlink { get; }
}
