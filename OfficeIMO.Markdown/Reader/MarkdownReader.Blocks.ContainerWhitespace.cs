namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>Expands only padding after a container marker, retaining its original tab stops.</summary>
    /// <remarks>Tabs inside the content remain literal; stripping container padding happens afterwards.</remarks>
    private static string ExpandWhitespaceAfterContainerMarker(string line, int markerEnd) {
        int end = markerEnd;
        int column = markerEnd;
        // One padding column and four code-indent columns are structural. Tabs after them are code payload.
        int structuralEndColumn = markerEnd + 5;
        while (end < line.Length && column < structuralEndColumn && (line[end] == ' ' || line[end] == '\t')) {
            column += line[end] == '\t' ? 4 - column % 4 : 1;
            end++;
        }
        return line.Substring(0, markerEnd) + new string(' ', column - markerEnd) + line.Substring(end);
    }
}
