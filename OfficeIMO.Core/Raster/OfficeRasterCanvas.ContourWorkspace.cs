using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Retain one small workspace per canvas. Complex paths can still use larger
    // transient buffers, but do not keep them alive with a long-lived canvas.
    private const int MaximumRetainedContourWorkspaceCapacity = 8192;
    private ContourCoverageWorkspace? _contourCoverageWorkspace;

    private ContourCoverageWorkspace TakeContourCoverageWorkspace() {
        ContourCoverageWorkspace? workspace = _contourCoverageWorkspace;
        _contourCoverageWorkspace = null;
        // Taking the workspace removes it from the cache, so nested paint cannot
        // overwrite scratch still in use by an outer contour operation.
        return workspace ?? new ContourCoverageWorkspace();
    }

    private void ReturnContourCoverageWorkspace(ContourCoverageWorkspace workspace) {
        workspace.Clear();
        if (_contourCoverageWorkspace == null && workspace.CanRetain) {
            _contourCoverageWorkspace = workspace;
        }
    }

    private sealed class ContourCoverageWorkspace {
        internal List<double> Boundaries { get; } = new List<double>();
        internal List<double> RowBoundaries { get; } = new List<double>();
        internal List<ContourCrossing> Crossings { get; } = new List<ContourCrossing>();
        internal List<ContourCrossing> RowCrossings { get; } = new List<ContourCrossing>();
        internal List<(double Weight, int Start, int Count)> Scanlines { get; } = new List<(double, int, int)>();

        internal bool CanRetain => Boundaries.Capacity <= MaximumRetainedContourWorkspaceCapacity
            && RowBoundaries.Capacity <= MaximumRetainedContourWorkspaceCapacity
            && Crossings.Capacity <= MaximumRetainedContourWorkspaceCapacity
            && RowCrossings.Capacity <= MaximumRetainedContourWorkspaceCapacity
            && Scanlines.Capacity <= MaximumRetainedContourWorkspaceCapacity;

        internal void Clear() {
            Boundaries.Clear();
            RowBoundaries.Clear();
            Crossings.Clear();
            RowCrossings.Clear();
            Scanlines.Clear();
        }
    }
}
