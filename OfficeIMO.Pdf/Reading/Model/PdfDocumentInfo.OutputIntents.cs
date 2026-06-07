namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    private IReadOnlyList<string>? _outputIntentSubtypes;
    private IReadOnlyList<string>? _outputConditionIdentifiers;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfOutputIntentInfo>>? _outputIntentsBySubtype;
    private IReadOnlyDictionary<string, IReadOnlyList<PdfOutputIntentInfo>>? _outputIntentsByOutputConditionIdentifier;

    /// <summary>Catalog output intent metadata discovered from /OutputIntents.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntents { get; }

    /// <summary>Number of catalog output intents discovered from /OutputIntents.</summary>
    public int OutputIntentCount => OutputIntents.Count;

    /// <summary>True when at least one catalog output intent was readable.</summary>
    public bool HasReadableOutputIntents => OutputIntentCount > 0;

    /// <summary>Distinct output intent subtypes in first-seen order.</summary>
    public IReadOnlyList<string> OutputIntentSubtypes {
        get {
            if (_outputIntentSubtypes is not null) {
                return _outputIntentSubtypes;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var subtypes = new List<string>();
            for (int i = 0; i < OutputIntents.Count; i++) {
                string? subtype = OutputIntents[i].Subtype;
                if (subtype is not null && seen.Add(subtype)) {
                    subtypes.Add(subtype);
                }
            }

            _outputIntentSubtypes = subtypes.AsReadOnly();
            return _outputIntentSubtypes;
        }
    }

    /// <summary>Distinct output condition identifiers in first-seen order.</summary>
    public IReadOnlyList<string> OutputConditionIdentifiers {
        get {
            if (_outputConditionIdentifiers is not null) {
                return _outputConditionIdentifiers;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var identifiers = new List<string>();
            for (int i = 0; i < OutputIntents.Count; i++) {
                string? identifier = OutputIntents[i].OutputConditionIdentifier;
                if (identifier is not null && seen.Add(identifier)) {
                    identifiers.Add(identifier);
                }
            }

            _outputConditionIdentifiers = identifiers.AsReadOnly();
            return _outputConditionIdentifiers;
        }
    }

    /// <summary>Output intents grouped by /S subtype.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfOutputIntentInfo>> OutputIntentsBySubtype {
        get {
            if (_outputIntentsBySubtype is not null) {
                return _outputIntentsBySubtype;
            }

            var grouped = new Dictionary<string, List<PdfOutputIntentInfo>>(StringComparer.Ordinal);
            for (int i = 0; i < OutputIntents.Count; i++) {
                AddOutputIntent(grouped, OutputIntents[i].Subtype, OutputIntents[i]);
            }

            _outputIntentsBySubtype = ToReadOnlyLookup(grouped);
            return _outputIntentsBySubtype;
        }
    }

    /// <summary>Output intents grouped by /OutputConditionIdentifier.</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<PdfOutputIntentInfo>> OutputIntentsByOutputConditionIdentifier {
        get {
            if (_outputIntentsByOutputConditionIdentifier is not null) {
                return _outputIntentsByOutputConditionIdentifier;
            }

            var grouped = new Dictionary<string, List<PdfOutputIntentInfo>>(StringComparer.Ordinal);
            for (int i = 0; i < OutputIntents.Count; i++) {
                AddOutputIntent(grouped, OutputIntents[i].OutputConditionIdentifier, OutputIntents[i]);
            }

            _outputIntentsByOutputConditionIdentifier = ToReadOnlyLookup(grouped);
            return _outputIntentsByOutputConditionIdentifier;
        }
    }

    /// <summary>Returns output intents with a matching /S subtype.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> GetOutputIntentsBySubtype(string subtype) {
        Guard.NotNullOrWhiteSpace(subtype, nameof(subtype));
        return OutputIntentsBySubtype.TryGetValue(subtype, out IReadOnlyList<PdfOutputIntentInfo>? intents)
            ? intents
            : Array.Empty<PdfOutputIntentInfo>();
    }

    /// <summary>Returns output intents with a matching /OutputConditionIdentifier.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> GetOutputIntentsByOutputConditionIdentifier(string outputConditionIdentifier) {
        Guard.NotNullOrWhiteSpace(outputConditionIdentifier, nameof(outputConditionIdentifier));
        return OutputIntentsByOutputConditionIdentifier.TryGetValue(outputConditionIdentifier, out IReadOnlyList<PdfOutputIntentInfo>? intents)
            ? intents
            : Array.Empty<PdfOutputIntentInfo>();
    }

    private static void AddOutputIntent(Dictionary<string, List<PdfOutputIntentInfo>> grouped, string? key, PdfOutputIntentInfo outputIntent) {
        if (key is null || key.Length == 0) {
            return;
        }

        if (!grouped.TryGetValue(key, out List<PdfOutputIntentInfo>? outputIntents)) {
            outputIntents = new List<PdfOutputIntentInfo>();
            grouped.Add(key, outputIntents);
        }

        outputIntents.Add(outputIntent);
    }
}
