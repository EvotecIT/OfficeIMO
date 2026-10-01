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
        builder.Append(tree.SourcePrefix);
        WriteGroup(tree.Root, builder, unicodeSkipCount: 1);
        builder.Append(tree.SourceSuffix);
        return builder.ToString();
    }

    private static void WriteGroup(RtfGroup group, StringBuilder builder, int unicodeSkipCount) {
        builder.Append('{');
        foreach (RtfNode child in group.Children) {
            WriteNode(child, builder, ref unicodeSkipCount);
        }

        builder.Append('}');
    }

    private static void WriteNode(RtfNode node, StringBuilder builder, ref int unicodeSkipCount) {
        switch (node) {
            case RtfGroup group:
                WriteGroup(group, builder, unicodeSkipCount);
                break;
            case RtfControlWord control:
                builder.Append(control.RawText);
                if (control.Name == "uc" && control.Parameter is >= 0) unicodeSkipCount = control.Parameter.Value;
                break;
            case RtfControlSymbol symbol:
                builder.Append(symbol.RawText);
                break;
            case RtfText text:
                bool needsOverride = text.EncodedUnicodeSkipCount.HasValue && text.EncodedUnicodeSkipCount != unicodeSkipCount;
                if (needsOverride) builder.Append(@"\uc").Append(text.EncodedUnicodeSkipCount!.Value.ToString(CultureInfo.InvariantCulture)).Append(' ');
                builder.Append(text.RawText);
                if (needsOverride) builder.Append(@"\uc").Append(unicodeSkipCount.ToString(CultureInfo.InvariantCulture)).Append(' ');
                break;
            case RtfBinary binary:
                builder.Append(binary.RawText);
                break;
        }
    }
}
