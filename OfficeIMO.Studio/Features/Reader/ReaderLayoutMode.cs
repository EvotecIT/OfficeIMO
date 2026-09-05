namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Defines how the active document's pages are arranged in the reader surface.</summary>
public enum ReaderLayoutMode {
    SinglePage,
    Continuous,
    TwoPage,
    Grid
}

/// <summary>Display metadata for one reader layout choice.</summary>
public sealed record ReaderLayoutChoice(ReaderLayoutMode Mode, string Label, string Description);
