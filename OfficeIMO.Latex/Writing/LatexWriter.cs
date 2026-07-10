namespace OfficeIMO.Latex;

/// <summary>LaTeX writer mode.</summary>
public enum LatexWriterMode {
    /// <summary>Retain original characters outside edited spans.</summary>
    Preserve = 0,
    /// <summary>Normalize line endings while retaining supported syntax.</summary>
    Canonical
}

/// <summary>Options for source-backed LaTeX writing.</summary>
public sealed class LatexWriterOptions {
    /// <summary>Writer mode.</summary>
    public LatexWriterMode Mode { get; set; } = LatexWriterMode.Preserve;
    /// <summary>Canonical line ending, or null to use the source preference.</summary>
    public string? LineEnding { get; set; }
}

/// <summary>Applies non-overlapping semantic source edits without reformatting untouched syntax.</summary>
public static class LatexWriter {
    /// <summary>Writes a document.</summary>
    public static string Write(LatexDocument document, LatexWriterOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new LatexWriterOptions();
        string lineEnding = options.LineEnding ?? document.Source.PreferredLineEnding;
        ValidateLineEnding(lineEnding);

        string output = document.Source.Text;
        ILatexSourceEdit[] edits = document.GetSourceEdits().Where(static edit => edit.IsModified)
            .GroupBy(static edit => new { Start = edit.EditSpan.Start.Offset, End = edit.EditSpan.End.Offset })
            .Select(static group => group.First())
            .OrderBy(static edit => edit.EditSpan.Start.Offset)
            .ToArray();
        ValidateNonOverlapping(edits);
        if (edits.Length > 0) {
            var builder = new StringBuilder(output.Length);
            int cursor = 0;
            for (int index = 0; index < edits.Length; index++) {
                ILatexSourceEdit edit = edits[index];
                builder.Append(output, cursor, edit.EditSpan.Start.Offset - cursor);
                builder.Append(edit.Replacement);
                cursor = edit.EditSpan.End.Offset;
            }
            builder.Append(output, cursor, output.Length - cursor);
            output = builder.ToString();
        }
        return options.Mode == LatexWriterMode.Canonical ? NormalizeLineEndings(output, lineEnding) : output;
    }

    private static void ValidateNonOverlapping(IReadOnlyList<ILatexSourceEdit> edits) {
        for (int index = 1; index < edits.Count; index++) {
            if (edits[index].EditSpan.Start.Offset < edits[index - 1].EditSpan.End.Offset) {
                throw new InvalidOperationException("Modified LaTeX semantic regions overlap. Edit either the container or its child, not both.");
            }
        }
    }

    private static string NormalizeLineEndings(string value, string lineEnding) {
        var output = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\r') {
                if (index + 1 < value.Length && value[index + 1] == '\n') index++;
                output.Append(lineEnding);
            } else if (value[index] == '\n') output.Append(lineEnding);
            else output.Append(value[index]);
        }
        return output.ToString();
    }

    private static void ValidateLineEnding(string value) {
        if (value != "\n" && value != "\r" && value != "\r\n") throw new ArgumentException("LineEnding must be LF, CR, or CRLF.");
    }
}
