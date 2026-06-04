using AngleSharp.Dom;
using System.Text;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ValidateDocumentLimits(IDocument document, HtmlToWordOptions options) {
            if (!options.MaxHtmlNodes.HasValue && !options.MaxHtmlDepth.HasValue) {
                return;
            }

            int nodeCount = 0;
            var stack = new Stack<(INode Node, int Depth)>();
            foreach (var child in document.ChildNodes) {
                stack.Push((child, 1));
            }

            while (stack.Count > 0) {
                var current = stack.Pop();
                nodeCount++;

                if (options.MaxHtmlNodes.HasValue && nodeCount > options.MaxHtmlNodes.Value) {
                    ThrowLimitExceeded(options, "HtmlNodeLimitExceeded", "HTML node count exceeded the configured conversion limit.", "MaxHtmlNodes", nodeCount, options.MaxHtmlNodes.Value);
                }

                if (options.MaxHtmlDepth.HasValue && current.Depth > options.MaxHtmlDepth.Value) {
                    ThrowLimitExceeded(options, "HtmlDepthLimitExceeded", "HTML nesting depth exceeded the configured conversion limit.", "MaxHtmlDepth", current.Depth, options.MaxHtmlDepth.Value);
                }

                for (int i = current.Node.ChildNodes.Length - 1; i >= 0; i--) {
                    stack.Push((current.Node.ChildNodes[i], current.Depth + 1));
                }
            }
        }

        private void ValidateCssLimit(string css, string? source) {
            if (!_options.MaxCssBytes.HasValue) {
                return;
            }

            var bytes = Encoding.UTF8.GetByteCount(css);
            if (bytes > _options.MaxCssBytes.Value) {
                ThrowLimitExceeded(_options, "CssSizeLimitExceeded", "CSS size exceeded the configured conversion limit.", source ?? "stylesheet", bytes, _options.MaxCssBytes.Value);
            }
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
            AddDiagnostic(options, code, message, source, new HtmlConversionLimitException(code, message, source, actual, limit, detail), HtmlConversionDiagnosticSeverity.Error);
            throw new HtmlConversionLimitException(code, message, source, actual, limit, detail);
        }
    }
}
