namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Writes an RTF syntax tree back to RTF while preserving raw tokens captured by the parser.
/// </summary>
public static class RtfSyntaxWriter {
    /// <summary>
    /// Serializes a syntax tree without applying semantic normalization.
    /// </summary>
    public static string Write(RtfSyntaxTree tree) {
        if (tree == null) throw new ArgumentNullException(nameof(tree));

        var builder = new StringBuilder();
        WriteGroup(tree.Root, builder);
        return builder.ToString();
    }

    private static void WriteGroup(RtfGroup group, StringBuilder builder) {
        builder.Append('{');
        foreach (RtfNode child in group.Children) {
            WriteNode(child, builder);
        }

        builder.Append('}');
    }

    private static void WriteNode(RtfNode node, StringBuilder builder) {
        switch (node) {
            case RtfGroup group:
                WriteGroup(group, builder);
                break;
            case RtfControlWord control:
                builder.Append(control.RawText);
                break;
            case RtfControlSymbol symbol:
                builder.Append(symbol.RawText);
                break;
            case RtfText text:
                builder.Append(text.RawText);
                break;
            case RtfBinary binary:
                builder.Append(binary.RawText);
                break;
        }
    }
}
