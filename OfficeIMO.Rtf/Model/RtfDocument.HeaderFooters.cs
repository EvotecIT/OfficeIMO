namespace OfficeIMO.Rtf;

/// <content>Resolves section header and footer inheritance.</content>
public sealed partial class RtfDocument {
    /// <summary>Returns the effective headers and footers for a section, including inherited kinds.</summary>
    /// <remarks>An explicit empty destination overrides inherited content. For repeated declarations, the last declaration of a kind wins.</remarks>
    public IReadOnlyList<RtfHeaderFooter> GetEffectiveHeaderFooters(RtfSection section) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (!_sections.Contains(section)) throw new ArgumentException("Section must belong to this document.", nameof(section));
        var owned = new HashSet<RtfHeaderFooter>(_sections.SelectMany(item => item.HeaderFooters));
        var effective = new Dictionary<RtfHeaderFooterKind, RtfHeaderFooter>();
        foreach (RtfHeaderFooter headerFooter in _headerFooters) {
            if (!owned.Contains(headerFooter)) effective[headerFooter.Kind] = headerFooter;
        }
        foreach (RtfSection current in _sections) {
            foreach (RtfHeaderFooter headerFooter in current.HeaderFooters) effective[headerFooter.Kind] = headerFooter;
            if (ReferenceEquals(current, section)) break;
        }
        return effective.OrderBy(item => item.Key).Select(item => item.Value).ToArray();
    }
}
