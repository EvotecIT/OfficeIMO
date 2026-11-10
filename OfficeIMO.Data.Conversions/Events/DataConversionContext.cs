using System.Diagnostics;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Shared state of one conversion: options, cancellation and the loss collector. Findings are aggregated per code,
/// with a count and the first locations, so a million-row file still produces a short report.
/// </summary>
internal sealed class DataConversionContext {
    private const int MaxLocations = 10;
    private readonly Dictionary<string, Finding> _findings = new(StringComparer.Ordinal);
    private readonly List<string> _order = new();
    private readonly Stopwatch _timer = Stopwatch.StartNew();

    public DataConversionContext(DataConversionOptions? options, CancellationToken cancellationToken = default) {
        Options = options ?? DataConversionOptions.Default;
        CancellationToken = cancellationToken;
    }

    public DataConversionOptions Options { get; }

    public CancellationToken CancellationToken { get; }

    /// <summary>The reader whose position locates findings.</summary>
    public DataEventReader? Source { get; set; }

    /// <summary>Items written: documents or table rows, as the target defines them.</summary>
    public long ItemsWritten { get; set; }

    private sealed class Finding {
        public Finding(string message, OfficeConversionLossKind loss) {
            Message = message;
            Loss = loss;
        }

        public string Message { get; }
        public OfficeConversionLossKind Loss { get; }
        public long Count;
        public readonly List<string> Locations = new();
    }

    /// <summary>Records one occurrence of a finding at the current source position.</summary>
    public void Report(string code, OfficeConversionLossKind loss, string message) => Report(code, loss, message, Source?.Location);

    public void Report(string code, OfficeConversionLossKind loss, string message, string? location) {
        if (!_findings.TryGetValue(code, out Finding? finding)) {
            finding = new Finding(message, loss);
            _findings.Add(code, finding);
            _order.Add(code);
        }

        finding.Count++;
        if (location != null && finding.Locations.Count < MaxLocations) finding.Locations.Add(location);
    }

    public bool HasLoss {
        get {
            foreach (Finding finding in _findings.Values) {
                if (finding.Loss != OfficeConversionLossKind.None) return true;
            }

            return false;
        }
    }

    public IReadOnlyList<OfficeDiagnostic> CreateDiagnostics(string format) {
        var diagnostics = new List<OfficeDiagnostic>(_order.Count);
        foreach (string code in _order) {
            Finding finding = _findings[code];
            string message = finding.Count == 1 ? finding.Message
                : finding.Message + " (" + finding.Count.ToString(CultureInfo.InvariantCulture) + " occurrences)";
            string? location = finding.Locations.Count == 0 ? null : string.Join("; ", finding.Locations);
            OfficeDiagnosticSeverity severity = finding.Loss == OfficeConversionLossKind.None
                ? OfficeDiagnosticSeverity.Information : OfficeDiagnosticSeverity.Warning;
            diagnostics.Add(new OfficeDiagnostic(code, severity, message, location, format, finding.Loss));
        }

        return diagnostics;
    }

    public OfficeConversionReport CreateReport(string source, string target, string? outputPath = null) =>
        new OfficeConversionReport(source, target, outputPath, ItemsWritten, _timer.Elapsed, CreateDiagnostics(source + " to " + target));

    /// <summary>Throws before publication when the caller asked to fail on loss.</summary>
    public void ThrowIfLossNotAllowed(string source, string target) {
        if (Options.FailOnLoss && HasLoss) CreateReport(source, target).ThrowIfLoss();
    }
}
