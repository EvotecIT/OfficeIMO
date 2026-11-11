namespace OfficeIMO.Yaml;

/// <summary>
/// The layout conventions of a loaded document, read once when the first edit needs them: mapping indent, sequence
/// indent and dash offset (the three numbers <see cref="YamlWriterOptions"/> uses), the dominant line break and the
/// preferred quote. New nodes follow them so an edit looks like the rest of the file.
/// </summary>
internal sealed class YamlLayoutProfile {
    /// <summary>Collections inspected at most; enough to read a file's conventions without walking a large stream.</summary>
    private const int SampleLimit = 512;

    public int MappingIndent { get; private set; } = 2;
    public int SequenceIndent { get; private set; } = 2;
    public int SequenceDashOffset { get; private set; }
    public string NewLine { get; private set; } = "\n";
    public YamlStyle Quote { get; private set; } = YamlStyle.SingleQuoted;

    public YamlWriterOptions ToWriterOptions() => new YamlWriterOptions {
        MappingIndent = MappingIndent,
        SequenceIndent = Math.Max(SequenceIndent, SequenceDashOffset + 2),
        SequenceDashOffset = SequenceDashOffset,
        NewLine = NewLine
    };

    public static YamlLayoutProfile Infer(YamlDocument document) {
        var profile = new YamlLayoutProfile();
        string source = document._source;
        profile.NewLine = DominantLineBreak(source);
        int mappingVote = 0, offsetVote = -1, sequenceVote = 0, singles = 0, doubles = 0, seen = 0;
        foreach (YamlDocumentPart part in document.Documents) {
            if (part.RootNode is null) continue;
            foreach (YamlNode node in part.RootNode.Descendants()) {
                if (seen > SampleLimit * 8) break;
                seen++;
                if (node is YamlScalar { HasSourceSpan: true } scalar) {
                    if (scalar._style == YamlStyle.SingleQuoted) singles++;
                    else if (scalar._style == YamlStyle.DoubleQuoted) doubles++;
                    continue;
                }
                // The first vote of each kind decides it; once both are in, the rest of the sample only counts quotes.
                if (mappingVote > 0 && offsetVote >= 0) continue;
                if (node is not YamlMapping { _style: YamlStyle.Block, HasSourceSpan: true } mapping || mapping._count == 0) continue;
                int keyColumn = Column(source, mapping.Entries[0].Key._start);
                for (int i = 0; i < mapping._count && (mappingVote == 0 || offsetVote < 0); i++) {
                    YamlNode value = mapping.Entries[i].Value;
                    if (!value.HasSourceSpan || value._style != YamlStyle.Block) continue;
                    if (value is YamlMapping { _count: > 0 } child && mappingVote == 0) {
                        int indent = Column(source, child.Entries[0].Key._start) - keyColumn;
                        if (indent > 0 && indent <= 9 && YamlTrivia.OnlyBlanksOrIndicators(source, YamlTrivia.LineStart(source, child.Entries[0].Key._start), child.Entries[0].Key._start)) mappingVote = indent;
                    } else if (value is YamlSequence { _count: > 0 } sequence && offsetVote < 0) {
                        int dash = YamlTrivia.DashPosition(source, sequence.Items[0]._start);
                        int offset = Column(source, dash) - keyColumn;
                        int content = Column(source, sequence.Items[0]._start) - keyColumn;
                        if (offset >= 0 && offset <= 8 && content >= offset + 2 && content <= 10) {
                            offsetVote = offset;
                            sequenceVote = content;
                        }
                    }
                }
            }
        }
        if (mappingVote > 0) profile.MappingIndent = mappingVote;
        if (offsetVote >= 0) {
            profile.SequenceDashOffset = offsetVote;
            profile.SequenceIndent = sequenceVote;
        } else {
            profile.SequenceIndent = 2;
        }
        if (doubles > singles) profile.Quote = YamlStyle.DoubleQuoted;
        return profile;
    }

    private static int Column(string source, int position) => position - YamlTrivia.LineStart(source, position);

    /// <summary>Returns <c>\r\n</c> when most line breaks are <c>\r\n</c>, otherwise <c>\n</c>.</summary>
    private static string DominantLineBreak(string source) {
        int carriage = source.IndexOf('\r');
        if (carriage < 0) return "\n";
        int crlf = 0, lf = 0;
        for (int i = 0; i < source.Length; i++) {
            if (source[i] != '\n') continue;
            if (i > 0 && source[i - 1] == '\r') crlf++;
            else lf++;
        }
        return crlf > lf ? "\r\n" : "\n";
    }
}
