namespace OfficeIMO.AsciiDoc;

/// <summary>Writes source-backed AsciiDoc documents.</summary>
internal static class AsciiDocWriter {
    /// <summary>Writes a document using preserve or canonical mode.</summary>
    internal static string Write(AsciiDocDocument document, AsciiDocWriterOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new AsciiDocWriterOptions();

        string lineEnding = options.LineEnding ?? document.Source.PreferredLineEnding;
        ValidateLineEnding(lineEnding);

        if (options.Mode == AsciiDocWriterMode.Preserve && !document.IsModified && document.SyntaxTree.IsLossless) {
            return document.Source.Text;
        }

        var context = new AsciiDocWriterContext(options.Mode, lineEnding);
        var builder = new StringBuilder(document.Source.Text.Length);
        for (int index = 0; index < document.Blocks.Count; index++) {
            AsciiDocBlock block = document.Blocks[index];
            if (document.IsStructureModified && index > 0) {
                AsciiDocBlock previous = document.Blocks[index - 1];
                bool originalNeighbors = ReferenceEquals(previous.Syntax.Parent, block.Syntax.Parent) && previous.Syntax.EndOffset == block.Syntax.StartOffset;
                if (!originalNeighbors && builder.Length > 0) {
                    if (builder[builder.Length - 1] != '\n' && builder[builder.Length - 1] != '\r') builder.Append(lineEnding);
                    if (previous is not AsciiDocBlankLine && block is not AsciiDocBlankLine &&
                        previous is not IAsciiDocBlockMetadata && previous is not AsciiDocListContinuation) builder.Append(lineEnding);
                }
            }
            builder.Append(block.Write(context));
        }
        return builder.ToString();
    }

    private static void ValidateLineEnding(string lineEnding) {
        if (lineEnding != "\n" && lineEnding != "\r" && lineEnding != "\r\n") {
            throw new ArgumentException("LineEnding must be LF, CR, or CRLF.", nameof(lineEnding));
        }
    }
}
