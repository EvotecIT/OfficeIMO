namespace OfficeIMO.Html;

internal sealed class HtmlRenderOperationBudget {
    internal long LayoutOperations { get; private set; }
    internal long BackgroundImageTiles { get; private set; }

    internal void Reset() {
        LayoutOperations = 0L;
        BackgroundImageTiles = 0L;
    }

    internal void ChargeLayoutOperations(long count, int limit, string source) {
        if (count < 0L || LayoutOperations > limit - count) {
            LayoutOperations = (long)limit + 1L;
        } else {
            LayoutOperations += count;
        }
        if (LayoutOperations > limit) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded,
                "HTML layout exceeded the configured operation limit at " + source + ".",
                nameof(HtmlRenderOptions.MaxLayoutOperations),
                LayoutOperations,
                limit);
        }
    }

    internal bool TryReserveBackgroundImageTiles(long count, int limit) {
        if (count <= 0L || BackgroundImageTiles > limit - count) return false;
        BackgroundImageTiles += count;
        return true;
    }
}
