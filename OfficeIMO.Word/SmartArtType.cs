namespace OfficeIMO.Word;

/// <summary>
/// SmartArt layout types supported by <see cref="WordDocument.AddSmartArt(SmartArtType)"/>.
/// </summary>
public enum SmartArtType {
    /// <summary>Basic process diagram.</summary>
    BasicProcess,
    /// <summary>Hierarchy layout diagram.</summary>
    Hierarchy,
    /// <summary>Cycle diagram layout.</summary>
    Cycle,
    /// <summary>Picture organization chart layout.</summary>
    PictureOrgChart,
    /// <summary>Continuous block process layout.</summary>
    ContinuousBlockProcess,
    /// <summary>Custom SmartArt template 1 (from exported design).</summary>
    CustomSmartArt1,
    /// <summary>Custom SmartArt template 2 (from exported design).</summary>
    CustomSmartArt2
}
