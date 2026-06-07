namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalDocument {
    /// <summary>Catalog output intent metadata discovered from /OutputIntents.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> OutputIntents { get; }

    /// <summary>Number of catalog output intents discovered from /OutputIntents.</summary>
    public int OutputIntentCount => OutputIntents.Count;

    /// <summary>True when at least one catalog output intent was readable.</summary>
    public bool HasReadableOutputIntents => OutputIntentCount > 0;

    /// <summary>Distinct output intent subtypes in first-seen order.</summary>
    public IReadOnlyList<string> OutputIntentSubtypes => OutputIntents.Select(intent => intent.Subtype).Where(subtype => subtype is not null).Cast<string>().Distinct(StringComparer.Ordinal).ToArray();

    /// <summary>Distinct output condition identifiers in first-seen order.</summary>
    public IReadOnlyList<string> OutputConditionIdentifiers => OutputIntents.Select(intent => intent.OutputConditionIdentifier).Where(identifier => identifier is not null).Cast<string>().Distinct(StringComparer.Ordinal).ToArray();

    /// <summary>Returns output intents with a matching /S subtype.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> GetOutputIntentsBySubtype(string subtype) {
        Guard.NotNullOrWhiteSpace(subtype, nameof(subtype));
        return OutputIntents.Where(intent => string.Equals(intent.Subtype, subtype, StringComparison.Ordinal)).ToArray();
    }

    /// <summary>Returns output intents with a matching /OutputConditionIdentifier.</summary>
    public IReadOnlyList<PdfOutputIntentInfo> GetOutputIntentsByOutputConditionIdentifier(string outputConditionIdentifier) {
        Guard.NotNullOrWhiteSpace(outputConditionIdentifier, nameof(outputConditionIdentifier));
        return OutputIntents.Where(intent => string.Equals(intent.OutputConditionIdentifier, outputConditionIdentifier, StringComparison.Ordinal)).ToArray();
    }
}
