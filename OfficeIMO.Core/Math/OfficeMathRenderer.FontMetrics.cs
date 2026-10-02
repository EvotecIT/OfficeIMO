using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeMathRenderer {
    private sealed partial class LayoutEngine {
        private LayoutBox Token(OfficeMathExpression expression, double scale) =>
            Text(_options.TokenPaintText?.Invoke(expression) ?? expression.Text ?? string.Empty,
                scale, expression.Text);

        private LayoutBox Text(string text, double scale, string? logicalText = null) {
            double size = FontSize(scale);
            OfficeTextMeasurementStyle style = _measurer.CreateStyle(_options.Font.WithSize(size), 72D);
            double width = Math.Max(size * 0.2D, _measurer.MeasureWidth(text, style));
            double advance = width;
            double height = size;
            // Positioned drawing text paints its baseline one source em below the frame top.
            double baseline = size;
            if (HasScopedTextCoverage(text, size)) {
                var ink = _measureScopedText(text, _options.Font.WithSize(size));
                advance = Math.Max(0.01D, ink.Advance);
                // Symmetric bearing room keeps centered positioned text at its natural
                // advance, while the frame contains overhangs and synthetic italic/bold.
                double bearing = Math.Max(0D, Math.Max(-ink.Left, ink.Right - advance));
                width = advance + bearing * 2D;
                if (ink.HasInk) {
                    baseline = Math.Max(0D, -ink.Top);
                    height = Math.Max(0.01D, ink.Bottom - ink.Top);
                }
            }
            var box = new LayoutBox(width, height, baseline);
            if (!string.IsNullOrEmpty(text)) {
                box.Commands.Add(LayoutCommand.TextCommand(text, 0D, 0D, width, height, size, baseline, advance, logicalText));
            }
            return box;
        }

        private bool HasScopedTextCoverage(string text, double size) {
            if (text.Length == 0) return false;
            foreach (OfficeFontFallbackRun run in _options.Fonts.PlanFallbackRuns(text,
                _options.Font.FamilyName, _options.Font.Style)) {
                if (!_options.Fonts.TryResolveFaceForText(run.Text, run.FamilyName,
                    _options.Font.Style, size, out _)) return false;
            }
            return true;
        }
    }
}
