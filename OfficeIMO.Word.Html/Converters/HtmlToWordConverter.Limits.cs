using AngleSharp.Dom;
using OfficeIMO.Html;
using System.Text;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ValidateDocumentLimits(IDocument document, HtmlToWordOptions options) {
            try {
                HtmlConversionInputGuard.ValidateDocument(document, options.Limits);
            } catch (HtmlDomLimitException exception) {
                ThrowLimitExceeded(
                    options,
                    exception.Code,
                    exception.Message,
                    exception.LimitSource,
                    exception.Actual,
                    exception.Limit);
            }
        }

        private void ValidateCssLimit(string css, string? source) {
            var bytes = Encoding.UTF8.GetByteCount(css);
            if (_options.MaxCssBytes.HasValue && bytes > _options.MaxCssBytes.Value) {
                ThrowLimitExceeded(_options, "CssSizeLimitExceeded", "CSS size exceeded the configured conversion limit.", source ?? "stylesheet", bytes, _options.MaxCssBytes.Value);
            }

            ReserveCssBytes(bytes, source ?? "stylesheet");
        }

        private void ReserveCssBytes(long length, string source) {
            if (!_options.MaxTotalCssBytes.HasValue) {
                return;
            }

            var remaining = _options.MaxTotalCssBytes.Value - _cssBytesUsed;
            if (length > remaining) {
                ThrowLimitExceeded(_options, "CssTotalSizeLimitExceeded", "Total CSS size exceeded the configured conversion limit.", source, _cssBytesUsed + length, _options.MaxTotalCssBytes.Value);
            }

            _cssBytesUsed += length;
        }

        private (long? Limit, bool LimitedByTotalBudget) GetCssReadLimit() {
            var limit = _options.MaxCssBytes;
            var limitedByTotalBudget = false;
            if (_options.MaxTotalCssBytes.HasValue) {
                var remaining = _options.MaxTotalCssBytes.Value - _cssBytesUsed;
                if (remaining <= 0) {
                    ThrowLimitExceeded(_options, "CssTotalSizeLimitExceeded", "Total CSS size exceeded the configured conversion limit.", "MaxTotalCssBytes", _cssBytesUsed, _options.MaxTotalCssBytes.Value);
                }

                if (!limit.HasValue || remaining < limit.Value) {
                    limit = remaining;
                    limitedByTotalBudget = true;
                }
            }

            return (limit, limitedByTotalBudget);
        }

        private void ValidateTableLimit(HtmlToWordOptions options, int rows, int columns) {
            if (!options.MaxTableCells.HasValue) {
                return;
            }

            var cells = (long)rows * columns;
            if (cells > options.MaxTableCells.Value) {
                ThrowLimitExceeded(options, "TableSizeLimitExceeded", "HTML table size exceeded the configured conversion limit.", "MaxTableCells", cells, options.MaxTableCells.Value);
            }
        }

        private void ThrowLimitExceeded(HtmlToWordOptions options, string code, string message, string source, long actual, long limit) {
            var detail = $"Actual={actual}; Limit={limit}";
            AddDiagnostic(options, code, message, source, new HtmlConversionLimitException(code, message, source, actual, limit, detail), HtmlDiagnosticSeverity.Error);
            throw new HtmlConversionLimitException(code, message, source, actual, limit, detail);
        }
    }
}
