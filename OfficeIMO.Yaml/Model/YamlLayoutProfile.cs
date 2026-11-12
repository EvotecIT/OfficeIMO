namespace OfficeIMO.Yaml;

/// <summary>
/// The layout conventions of a loaded document, read from its node table once when the first edit needs them: mapping indent, sequence
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
            // The loaded nodes are read from the node table, so inferring the layout creates no node objects.
            if (part._table is not { } table || part._rootRecord < 0) continue;
            YamlNodeRecord[] records = table.Records;
            int end = part._rootRecord + records[part._rootRecord].Size;
            for (int node = part._rootRecord; node < end; node++) {
                if (seen > SampleLimit * 8) break;
                seen++;
                ref YamlNodeRecord record = ref records[node];
                if (record.Kind == YamlRecordKind.Scalar) {
                    if (record.Style == YamlStyle.SingleQuoted) singles++;
                    else if (record.Style == YamlStyle.DoubleQuoted) doubles++;
                    continue;
                }
                // The first vote of each kind decides it; once both are in, the rest of the sample only counts quotes.
                if (mappingVote > 0 && offsetVote >= 0) continue;
                if (record.Kind != YamlRecordKind.Mapping || record.Style != YamlStyle.Block || record.Count == 0) continue;
                int key = node + 1;
                int keyColumn = Column(source, records[key].Start);
                for (int i = 0; i < record.Count && (mappingVote == 0 || offsetVote < 0); i++, key = table.Next(table.Next(key))) {
                    int value = table.Next(key);
                    ref YamlNodeRecord child = ref records[value];
                    if (child.Style != YamlStyle.Block || child.Count == 0) continue;
                    if (child.Kind == YamlRecordKind.Mapping && mappingVote == 0) {
                        int childKey = records[value + 1].Start;
                        int indent = Column(source, childKey) - keyColumn;
                        if (indent > 0 && indent <= 9 && YamlTrivia.OnlyBlanksOrIndicators(source, YamlTrivia.LineStart(source, childKey), childKey)) mappingVote = indent;
                    } else if (child.Kind == YamlRecordKind.Sequence && offsetVote < 0) {
                        int item = records[value + 1].Start;
                        int dash = YamlTrivia.DashPosition(source, item);
                        int offset = Column(source, dash) - keyColumn;
                        int content = Column(source, item) - keyColumn;
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
