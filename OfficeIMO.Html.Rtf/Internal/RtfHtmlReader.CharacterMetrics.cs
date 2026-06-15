namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyCharacterMetrics(RtfRun run) {
            run.CharacterSpacingTwips = ResolveResettableInt(style => style.CharacterSpacingTwips, resetValue: 0);
            run.CharacterScalePercent = ResolveResettableInt(style => style.CharacterScalePercent, resetValue: 100);
            run.CharacterOffsetHalfPoints = ResolveResettableInt(style => style.CharacterOffsetHalfPoints, resetValue: 0);
        }

        private int? ResolveResettableInt(Func<HtmlStyleDeclaration, int?> selector, int resetValue) {
            foreach (HtmlStyleScope scope in _styles) {
                int? value = selector(scope.Style);
                if (value.HasValue) {
                    return value.Value == resetValue ? null : value.Value;
                }
            }

            return null;
        }
    }
}
