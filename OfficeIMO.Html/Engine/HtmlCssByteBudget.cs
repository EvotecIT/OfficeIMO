using System.Text;

namespace OfficeIMO.Html;

/// <summary>Owns per-stylesheet and operation-wide CSS byte accounting.</summary>
internal sealed class HtmlCssByteBudget {
    private readonly HtmlConversionLimits _limits;
    private long _totalBytes;

    internal HtmlCssByteBudget(HtmlConversionLimits limits) {
        _limits = (limits ?? throw new ArgumentNullException(nameof(limits))).Clone();
    }

    internal bool TryReserve(string css, out HtmlDomLimitException? exception) {
        if (!_limits.MaxCssBytes.HasValue && !_limits.MaxTotalCssBytes.HasValue) {
            exception = null;
            return true;
        }

        long bytes = Encoding.UTF8.GetByteCount(css ?? string.Empty);
        if (_limits.MaxCssBytes.HasValue && bytes > _limits.MaxCssBytes.Value) {
            exception = CreateException(
                HtmlConversionDiagnosticCodes.CssSizeLimitExceeded,
                nameof(HtmlConversionLimits.MaxCssBytes),
                bytes,
                _limits.MaxCssBytes.Value);
            return false;
        }

        long nextTotal = _totalBytes + bytes;
        if (_limits.MaxTotalCssBytes.HasValue && nextTotal > _limits.MaxTotalCssBytes.Value) {
            exception = CreateException(
                HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded,
                nameof(HtmlConversionLimits.MaxTotalCssBytes),
                nextTotal,
                _limits.MaxTotalCssBytes.Value);
            return false;
        }

        _totalBytes = nextTotal;
        exception = null;
        return true;
    }

    internal void ReserveOrThrow(string css) {
        if (!TryReserve(css, out HtmlDomLimitException? exception)) throw exception!;
    }

    private static HtmlDomLimitException CreateException(string code, string source, long actual, long limit) =>
        new HtmlDomLimitException(code, "CSS exceeded the configured conversion limit.", source, actual, limit);
}
