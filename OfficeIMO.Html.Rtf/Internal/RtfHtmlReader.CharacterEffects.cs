namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyCharacterEffects(RtfRun run) {
            run.Hidden = ResolveStyleValue(style => style.Hidden, fallback: false);
            run.Outline = ResolveStyleValue(style => style.Outline, fallback: false);
            run.Shadow = ResolveStyleValue(style => style.Shadow, fallback: false);
            run.Emboss = ResolveStyleValue(style => style.Emboss, fallback: false);
            run.Imprint = ResolveStyleValue(style => style.Imprint, fallback: false);
        }
    }
}
