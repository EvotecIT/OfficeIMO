namespace OfficeIMO.Pdf;

/// <summary>Captures the final page regions occupied by one or more flow groups.</summary>
public sealed class PdfLayoutPositionCapture {
    private readonly object _sync = new object();
    private readonly List<PdfLayoutRegion> _regions = new List<PdfLayoutRegion>();
    private bool _wasSkipped;

    /// <summary>Captured regions from the most recent layout pass.</summary>
    public IReadOnlyList<PdfLayoutRegion> Regions {
        get {
            lock (_sync) {
                return _regions.ToArray();
            }
        }
    }

    /// <summary>Last captured region, or null when the group emitted no content.</summary>
    public PdfLayoutRegion? Last {
        get {
            lock (_sync) {
                return _regions.Count == 0 ? null : _regions[_regions.Count - 1];
            }
        }
    }

    /// <summary>True when the associated group was skipped by its condition or overflow policy.</summary>
    public bool WasSkipped {
        get {
            lock (_sync) {
                return _wasSkipped;
            }
        }
    }

    internal void BeginLayoutPass() {
        lock (_sync) {
            _regions.Clear();
            _wasSkipped = false;
        }
    }

    internal void Add(PdfLayoutRegion region) {
        lock (_sync) {
            _regions.Add(region);
        }
    }

    internal void MarkSkipped() {
        lock (_sync) {
            _wasSkipped = true;
        }
    }
}
