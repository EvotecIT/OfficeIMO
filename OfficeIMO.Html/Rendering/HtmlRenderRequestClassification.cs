namespace OfficeIMO.Html;

/// <summary>Identifies which kind of immutable document snapshot is supplied to a render operation.</summary>
public enum HtmlRenderDocumentState {
    /// <summary>The inert source snapshot produced by parsing or loading HTML.</summary>
    StaticSource,
    /// <summary>An OfficeIMO-owned snapshot produced after explicit DOM edits.</summary>
    EditedSnapshot,
    /// <summary>A frozen document imported from another capture provider.</summary>
    ImportedCapture,
    /// <summary>A frozen snapshot captured from an OfficeIMO HTML runtime session.</summary>
    RuntimeSnapshot
}

/// <summary>Selects the geometry model used to contain rendered HTML.</summary>
public enum HtmlRenderLayoutSurface {
    /// <summary>One fixed-width and fixed-height viewport clipped to its bounds.</summary>
    Viewport,
    /// <summary>One fixed-width surface whose height grows with document content.</summary>
    Continuous,
    /// <summary>One or more fixed-size page sheets.</summary>
    Paged
}

/// <summary>Selects how content becomes one or more output surfaces.</summary>
public enum HtmlRenderPaginationPolicy {
    /// <summary>No pagination; content remains on one viewport or continuous surface.</summary>
    None,
    /// <summary>Layout reflows and fragments content into page sheets.</summary>
    FragmentedReflow,
    /// <summary>One completed continuous layout is sliced into fixed-size page canvases.</summary>
    FixedCanvasSlicing,
    /// <summary>Whole elements are placed onto pages without normal CSS fragmentation.</summary>
    ElementAwarePlacement
}

/// <summary>Identifies the output adapter intended to consume the retained render result.</summary>
public enum HtmlRenderEncoder {
    /// <summary>Retain the backend-neutral display list.</summary>
    DisplayList,
    /// <summary>Encode one or more PNG images.</summary>
    Png,
    /// <summary>Encode one or more JPEG images.</summary>
    Jpeg,
    /// <summary>Encode one or more TIFF images.</summary>
    Tiff,
    /// <summary>Encode one or more WebP images.</summary>
    Webp,
    /// <summary>Encode one or more SVG documents.</summary>
    Svg,
    /// <summary>Encode a PDF document.</summary>
    Pdf,
    /// <summary>Consume the retained geometry and source identity.</summary>
    Geometry,
    /// <summary>Consume the retained geometry through hit testing.</summary>
    HitTest
}

/// <summary>Identifies a built-in combination of CSS, layout, and pagination behavior.</summary>
public enum HtmlRenderIntentProfile {
    /// <summary>Screen CSS in one exact clipped viewport.</summary>
    ScreenViewport,
    /// <summary>Screen CSS in one content-height continuous surface.</summary>
    ScreenFullPage,
    /// <summary>Print CSS reflowed and fragmented into page sheets.</summary>
    PrintPaged,
    /// <summary>Screen CSS recomputed in a paged layout environment.</summary>
    ScreenMediaPaged,
    /// <summary>One screen layout frozen before fixed-canvas page slicing.</summary>
    ScreenSnapshotPaged,
    /// <summary>A continuous vector-oriented display-list surface using explicit CSS media.</summary>
    ContinuousVector
}

/// <summary>Selects how pages from a retained render are exposed to an output adapter.</summary>
public enum HtmlRenderPageSetMode {
    /// <summary>Keep every page as a separate ordered surface.</summary>
    Separate,
    /// <summary>Keep one selected zero-based page.</summary>
    Selected,
    /// <summary>Keep a bounded range of pages as separate ordered surfaces.</summary>
    Range,
    /// <summary>Place all selected pages vertically on one continuous surface.</summary>
    Stitched,
    /// <summary>Package separate encoded pages together with a manifest.</summary>
    ArchiveWithManifest
}
