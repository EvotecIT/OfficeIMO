using System;
using System.Threading;

namespace OfficeIMO.Drawing;

// Native document formats distinguish the caps of dash pieces and open figures.
// These options keep that geometry in Core without adding format-specific shape APIs.
internal enum OfficeStrokeOutlineCap { Flat, Round, Square, Triangle }

internal sealed class OfficeStrokeOutlineOptions {
    internal OfficeStrokeOutlineCap StartCap { get; set; }
    internal OfficeStrokeOutlineCap EndCap { get; set; }
    internal OfficeStrokeOutlineCap DashCap { get; set; }
    internal bool ClipMiter { get; set; }
    internal bool DegenerateMiterLimitOne { get; set; }
    // Native figures carry explicit closure and must retain geometry before its transform.
    internal bool? Closed { get; set; }
    internal bool PreserveExactGeometry { get; set; }
    internal CancellationToken CancellationToken { get; set; }
    internal Action<int>? ChargePoints { get; set; }

    internal static OfficeStrokeOutlineCap FromCap(OfficeStrokeLineCap cap) =>
        cap == OfficeStrokeLineCap.Round ? OfficeStrokeOutlineCap.Round :
        cap == OfficeStrokeLineCap.Square ? OfficeStrokeOutlineCap.Square : OfficeStrokeOutlineCap.Flat;
}
