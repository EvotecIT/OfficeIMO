namespace OfficeIMO.Pdf;

internal sealed class PdfTextAppearanceSegment {
    public PdfTextAppearanceSegment(string fontResourceName, PdfTextShowCommand command) {
        Guard.NotNullOrWhiteSpace(fontResourceName, nameof(fontResourceName));
        Guard.NotNull(command, nameof(command));

        FontResourceName = fontResourceName;
        Command = command;
    }

    public string FontResourceName { get; }

    public PdfTextShowCommand Command { get; }
}
