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
            builder.Append(document.Blocks[index].Write(context));
        }
        return builder.ToString();
    }

    private static void ValidateLineEnding(string lineEnding) {
        if (lineEnding != "\n" && lineEnding != "\r" && lineEnding != "\r\n") {
            throw new ArgumentException("LineEnding must be LF, CR, or CRLF.", nameof(lineEnding));
        }
    }
}
