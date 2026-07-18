using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Write;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>Analyzes whether this presentation fits the native binary PPT/POT/PPS writer subset.</summary>
        public LegacyPptWritePreflightReport AnalyzeLegacyPptWrite() {
            ThrowIfDisposed();
            return LegacyPptWritePreflight.Analyze(this);
        }
    }
}
