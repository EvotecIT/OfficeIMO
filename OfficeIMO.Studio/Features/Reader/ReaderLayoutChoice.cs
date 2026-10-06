namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Display metadata for one reader layout choice.</summary>
public sealed record ReaderLayoutChoice(ReaderLayoutMode Mode, string Label, string Description) {
    /// <summary>The readable selection value exposed by native accessibility.</summary>
    public override string ToString() => Label;
}
