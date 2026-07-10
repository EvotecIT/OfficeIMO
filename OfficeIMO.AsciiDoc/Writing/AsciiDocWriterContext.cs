namespace OfficeIMO.AsciiDoc;

internal sealed class AsciiDocWriterContext {
    internal AsciiDocWriterContext(AsciiDocWriterMode mode, string lineEnding) {
        Mode = mode;
        LineEnding = lineEnding;
    }

    internal AsciiDocWriterMode Mode { get; }
    internal string LineEnding { get; }
}
