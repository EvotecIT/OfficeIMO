namespace OfficeIMO.Drawing;

/// <summary>Stage reported by a single or batch image-export operation.</summary>
public enum OfficeImageExportProgressStage {
    /// <summary>An image is being rendered.</summary>
    Rendering,
    /// <summary>An encoded image is being committed to its destination.</summary>
    Saving,
    /// <summary>An image completed rendering or saving.</summary>
    Completed
}

/// <summary>Progress snapshot for a single or batch image-export operation.</summary>
public sealed class OfficeImageExportProgress {
    /// <summary>Creates an immutable progress snapshot.</summary>
    public OfficeImageExportProgress(
        OfficeImageExportProgressStage stage,
        int completedCount,
        int? totalCount = null,
        string? name = null,
        string? destinationPath = null) {
        Stage = stage;
        CompletedCount = completedCount;
        TotalCount = totalCount;
        Name = name;
        DestinationPath = destinationPath;
    }

    /// <summary>Current export stage.</summary>
    public OfficeImageExportProgressStage Stage { get; }

    /// <summary>Number of completed results.</summary>
    public int CompletedCount { get; }

    /// <summary>Total expected result count when known.</summary>
    public int? TotalCount { get; }

    /// <summary>Current result name when known.</summary>
    public string? Name { get; }

    /// <summary>Current normalized destination path when saving.</summary>
    public string? DestinationPath { get; }
}
