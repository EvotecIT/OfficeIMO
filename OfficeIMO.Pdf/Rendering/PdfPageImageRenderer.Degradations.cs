using System.Collections;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageImageRenderer {
    /// <summary>
    /// Adds the content a render skipped while drawing to the page's pre-scanned capability diagnostics, within the
    /// per-page diagnostic count.
    /// </summary>
    private static IReadOnlyList<PdfRenderCapabilityDiagnostic> AppendRenderDegradations(
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilityDiagnostics,
        IReadOnlyList<PdfRenderCapabilityDiagnostic> renderDegradations,
        bool rasterElementOmitted,
        int maximumDiagnostics) {
        if (renderDegradations.Count == 0 && !rasterElementOmitted) return capabilityDiagnostics;
        var merged = new List<PdfRenderCapabilityDiagnostic>(capabilityDiagnostics);
        foreach (PdfRenderCapabilityDiagnostic degradation in renderDegradations) {
            if (merged.Count >= maximumDiagnostics) return merged;
            if (!merged.Any(existing => existing.Code == degradation.Code && existing.Subject == degradation.Subject)) merged.Add(degradation);
        }
        if (rasterElementOmitted && merged.Count < maximumDiagnostics) {
            merged.Add(new PdfRenderCapabilityDiagnostic(PdfRenderCapabilities.Get(PdfRenderCapabilities.RasterWorkLimitId)));
        }
        return merged;
    }

    /// <summary>
    /// Raster diagnostic sink that only remembers whether an element was omitted at the raster work limit, so a page
    /// with many such elements, or many unrelated raster diagnostics, does not retain them all.
    /// </summary>
    private sealed class RasterOmissionSink : ICollection<OfficeImageExportDiagnostic> {
        internal bool ElementOmitted { get; private set; }

        public int Count => 0;
        public bool IsReadOnly => false;

        public void Add(OfficeImageExportDiagnostic item) {
            if (item?.Code == OfficeRasterWorkLimitException.DiagnosticCode) ElementOmitted = true;
        }

        public void Clear() => ElementOmitted = false;
        public bool Contains(OfficeImageExportDiagnostic item) => false;
        public void CopyTo(OfficeImageExportDiagnostic[] array, int arrayIndex) { }
        public bool Remove(OfficeImageExportDiagnostic item) => false;
        public IEnumerator<OfficeImageExportDiagnostic> GetEnumerator() => Enumerable.Empty<OfficeImageExportDiagnostic>().GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
