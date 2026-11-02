using System.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    private OfficeDiagnostic[] _sourceLoadDiagnostics = Array.Empty<OfficeDiagnostic>();
    /// <summary>Loads a presentation and returns findings from its source projection.</summary>
    public static PowerPointDocument Load(string path, PowerPointLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        var value = Load(path, options);
        report = value.CreateLoadReport(path, timer.Elapsed);
        return value;
    }
    /// <summary>Loads a presentation stream and returns findings from its source projection.</summary>
    public static PowerPointDocument Load(Stream stream, PowerPointLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        var value = Load(stream, options);
        report = value.CreateLoadReport("PowerPoint stream", timer.Elapsed);
        return value;
    }
    internal OfficeConversionReport CreateLoadReport(string source, TimeSpan elapsed) =>
        new OfficeConversionReport(source, "PowerPoint", itemsWritten: Slides.Count, elapsed: elapsed, diagnostics: _sourceLoadDiagnostics);
}
