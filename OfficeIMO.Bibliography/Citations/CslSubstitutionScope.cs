namespace OfficeIMO.Bibliography;

/// <summary>Retains rendering state while one candidate for a name substitution is evaluated.</summary>
internal sealed class CslSubstitutionScope {
    internal CslSubstitutionScope(CslSubstitutionScope? parent, CslContext context) {
        Parent = parent;
        PreviouslyRendered = new HashSet<string>(context.RenderedVariables, StringComparer.Ordinal);
        PreviouslySuppressed = new HashSet<string>(context.Suppressed, StringComparer.Ordinal);
        FirstNamesHandled = context.FirstNamesHandled;
        FirstNames = context.FirstNames;
        FirstNamesText = context.FirstNamesText;
        CompleteNamesMatch = context.CompleteNamesMatch;
        EmptyNamesReplacementCount = context.EmptyNamesReplacementCount;
        ObservedNamesCount = context.ObservedNames.Count;
        NarrativeNames = context.NarrativeNames;
        YearRendered = context.YearRendered;
    }
    internal CslSubstitutionScope? Parent { get; }
    internal ISet<string> PreviouslyRendered { get; }
    internal ISet<string> PreviouslySuppressed { get; }
    internal bool FirstNamesHandled { get; }
    internal string[] FirstNames { get; }
    internal string FirstNamesText { get; }
    internal bool CompleteNamesMatch { get; }
    internal int EmptyNamesReplacementCount { get; }
    internal int ObservedNamesCount { get; }
    internal CslText? NarrativeNames { get; }
    internal bool YearRendered { get; }
}
