namespace OfficeIMO.IWork.Internal;

/// <summary>Describes why recovered text is incomplete without confusing valid object markers with encoding errors.</summary>
internal static class IWorkTextDiagnostics {
    internal static string Describe(IWorkTextContent? content) {
        if (content == null) return "The text storage is malformed.";
        var reasons = new List<string>();
        if (content.HasInvalidSourceText) reasons.Add("invalid source text or text fields");
        if (content.HasUnresolvedInlineObjects) reasons.Add("unresolved inline objects");
        if (!content.IsFormattingComplete) reasons.Add("unsupported or unresolved formatting");
        return reasons.Count == 0 ? "The text storage is incomplete."
            : "The text storage contains " + string.Join(", ", reasons) + ".";
    }
}
