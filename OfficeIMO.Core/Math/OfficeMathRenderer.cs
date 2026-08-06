using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Settings for dependency-free mathematical layout and drawing.</summary>
public sealed class OfficeMathRenderOptions {
    /// <summary>Base mathematical font.</summary>
    public OfficeFontInfo Font { get; set; } = new OfficeFontInfo("Cambria Math", 18D);

    /// <summary>Foreground color.</summary>
    public OfficeColor Color { get; set; } = OfficeColor.Black;

    /// <summary>Optional canvas background.</summary>
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>Padding around the expression in drawing units.</summary>
    public double Padding { get; set; } = 8D;

    /// <summary>Relative scale applied to scripts, limits, and root indices.</summary>
    public double ScriptScale { get; set; } = 0.7D;

    /// <summary>Gap around fraction and decoration rules.</summary>
    public double RuleGap { get; set; } = 2D;

    /// <summary>Thickness of fraction, radical, bar, and box rules.</summary>
    public double RuleThickness { get; set; } = 1D;

    /// <summary>Horizontal and vertical gap between matrix cells.</summary>
    public double MatrixGap { get; set; } = 8D;

    /// <summary>DPI used by deterministic text measurement.</summary>
    public double Dpi { get; set; } = OfficeTextMeasurer.DefaultDpi;

    /// <summary>Creates a detached copy.</summary>
    public OfficeMathRenderOptions Clone() => new OfficeMathRenderOptions {
        Font = Font,
        Color = Color,
        BackgroundColor = BackgroundColor,
        Padding = Padding,
        ScriptScale = ScriptScale,
        RuleGap = RuleGap,
        RuleThickness = RuleThickness,
        MatrixGap = MatrixGap,
        Dpi = Dpi
    };

    internal void Validate() {
        Positive(Font.Size, nameof(Font));
        NonNegative(Padding, nameof(Padding));
        Positive(ScriptScale, nameof(ScriptScale));
        NonNegative(RuleGap, nameof(RuleGap));
        Positive(RuleThickness, nameof(RuleThickness));
        NonNegative(MatrixGap, nameof(MatrixGap));
        Positive(Dpi, nameof(Dpi));
    }

    private static void Positive(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) throw new ArgumentOutOfRangeException(name);
    }

    private static void NonNegative(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) throw new ArgumentOutOfRangeException(name);
    }
}

/// <summary>Measured bounds and baseline of a mathematical expression.</summary>
public readonly struct OfficeMathLayoutMetrics {
    internal OfficeMathLayoutMetrics(double width, double height, double baseline) {
        Width = width;
        Height = height;
        Baseline = baseline;
    }

    /// <summary>Expression width excluding caller-owned padding.</summary>
    public double Width { get; }

    /// <summary>Expression height excluding caller-owned padding.</summary>
    public double Height { get; }

    /// <summary>Distance from the top edge to the mathematical baseline.</summary>
    public double Baseline { get; }
}

/// <summary>Renders the shared mathematical expression tree into the shared drawing scene.</summary>
public static class OfficeMathRenderer {
    /// <summary>Measures an expression using deterministic, dependency-free font metrics.</summary>
    public static OfficeMathLayoutMetrics Measure(OfficeMathExpression expression, OfficeMathRenderOptions? options = null) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        OfficeMathRenderOptions effective = options?.Clone() ?? new OfficeMathRenderOptions();
        effective.Validate();
        LayoutBox box = new LayoutEngine(effective).Layout(expression, 1D);
        return new OfficeMathLayoutMetrics(box.Width, box.Height, box.Baseline);
    }

    /// <summary>Creates a tightly sized drawing containing an expression.</summary>
    public static OfficeDrawing Render(OfficeMathExpression expression, OfficeMathRenderOptions? options = null) {
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        OfficeMathRenderOptions effective = options?.Clone() ?? new OfficeMathRenderOptions();
        effective.Validate();
        LayoutBox box = new LayoutEngine(effective).Layout(expression, 1D);
        double width = Math.Max(1D, box.Width + effective.Padding * 2D);
        double height = Math.Max(1D, box.Height + effective.Padding * 2D);
        var drawing = new OfficeDrawing(width, height);
        if (effective.BackgroundColor.HasValue) {
            OfficeShape background = OfficeShape.Rectangle(width, height);
            background.FillColor = effective.BackgroundColor.Value;
            background.StrokeWidth = 0D;
            drawing.AddShape(background, 0D, 0D);
        }
        Paint(drawing, box, effective.Padding, effective.Padding, effective);
        return drawing;
    }

    /// <summary>Adds a laid-out expression to an existing drawing.</summary>
    public static OfficeMathLayoutMetrics AddToDrawing(
        OfficeDrawing drawing,
        OfficeMathExpression expression,
        double x,
        double y,
        OfficeMathRenderOptions? options = null) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (expression == null) throw new ArgumentNullException(nameof(expression));
        OfficeMathRenderOptions effective = options?.Clone() ?? new OfficeMathRenderOptions();
        effective.Validate();
        LayoutBox box = new LayoutEngine(effective).Layout(expression, 1D);
        if (x < 0D || y < 0D || x + box.Width > drawing.Width || y + box.Height > drawing.Height) {
            throw new ArgumentOutOfRangeException(nameof(expression), "The mathematical expression must fit inside the drawing bounds.");
        }
        Paint(drawing, box, x, y, effective);
        return new OfficeMathLayoutMetrics(box.Width, box.Height, box.Baseline);
    }

    private static void Paint(OfficeDrawing drawing, LayoutBox box, double x, double y, OfficeMathRenderOptions options) {
        for (int index = 0; index < box.Commands.Count; index++) {
            LayoutCommand command = box.Commands[index];
            if (command.Kind == LayoutCommandKind.Text) {
                if (string.IsNullOrEmpty(command.Text)) continue;
                drawing.AddText(command.Text!, x + command.X, y + command.Y, Math.Max(0.01D, command.Width),
                    Math.Max(0.01D, command.Height), options.Font.WithSize(command.FontSize), options.Color,
                    OfficeTextAlignment.Left, command.Height);
            } else if (command.Kind == LayoutCommandKind.Line) {
                OfficeShape line = OfficeShape.Line(x + command.X, y + command.Y, x + command.X2, y + command.Y2);
                line.StrokeColor = options.Color;
                line.StrokeWidth = options.RuleThickness;
                drawing.AddShape(line, Math.Min(x + command.X, x + command.X2), Math.Min(y + command.Y, y + command.Y2));
            } else {
                OfficeShape rectangle = OfficeShape.Rectangle(command.Width, command.Height);
                rectangle.FillColor = OfficeColor.Transparent;
                rectangle.StrokeColor = options.Color;
                rectangle.StrokeWidth = options.RuleThickness;
                drawing.AddShape(rectangle, x + command.X, y + command.Y);
            }
        }
    }

    private sealed class LayoutEngine {
        private readonly OfficeMathRenderOptions _options;
        private readonly OfficeTextMeasurer _measurer;

        internal LayoutEngine(OfficeMathRenderOptions options) {
            _options = options;
            _measurer = OfficeTextMeasurer.Create(options.Font);
        }

        internal LayoutBox Layout(OfficeMathExpression expression, double scale) {
            switch (expression.Kind) {
                case OfficeMathKind.Text:
                case OfficeMathKind.Identifier:
                case OfficeMathKind.Number:
                case OfficeMathKind.Operator:
                    return Text(expression.Text ?? string.Empty, scale);
                case OfficeMathKind.Row: return Row(expression.Children, scale);
                case OfficeMathKind.Fraction: return Fraction(expression, scale);
                case OfficeMathKind.Radical: return Radical(expression, scale);
                case OfficeMathKind.Superscript: return Scripts(expression, scale, false, true);
                case OfficeMathKind.Subscript: return Scripts(expression, scale, true, false);
                case OfficeMathKind.SubSuperscript: return Scripts(expression, scale, true, true);
                case OfficeMathKind.LeftSubSuperscript: return LeftScripts(expression, scale);
                case OfficeMathKind.LowerLimit: return Limit(expression, scale, over: false);
                case OfficeMathKind.UpperLimit: return Limit(expression, scale, over: true);
                case OfficeMathKind.SlashedFraction: return SlashedFraction(expression, scale);
                case OfficeMathKind.Nary: return Nary(expression, scale);
                case OfficeMathKind.Delimited: return Delimited(expression, scale);
                case OfficeMathKind.DelimiterList: return DelimiterList(expression, scale);
                case OfficeMathKind.Function:
                    return Row(new[] { OfficeMath.Identifier(expression.Text ?? string.Empty), OfficeMath.Delimited(expression.Children[0]) }, scale);
                case OfficeMathKind.Matrix:
                case OfficeMathKind.EquationArray:
                    return Matrix(expression, scale);
                case OfficeMathKind.Accent: return Accent(expression, scale);
                case OfficeMathKind.Overbar: return Bar(expression.Children[0], scale, true);
                case OfficeMathKind.Underbar: return Bar(expression.Children[0], scale, false);
                case OfficeMathKind.Box: return Box(expression.Children[0], scale);
                case OfficeMathKind.Phantom:
                    LayoutBox phantom = Layout(expression.Children[0], scale);
                    phantom.Commands.Clear();
                    return phantom;
                case OfficeMathKind.Stack:
                case OfficeMathKind.StretchStack:
                    return Stack(expression.Children, scale);
                default: return Text(expression.ToPlainText(), scale);
            }
        }

        private LayoutBox Text(string text, double scale) {
            double fontSize = _options.Font.Size * scale;
            OfficeTextMeasurementStyle style = _measurer.CreateStyle(_options.Font.WithSize(fontSize), _options.Dpi);
            double width = Math.Max(fontSize * 0.2D, _measurer.MeasureWidth(text, style));
            double height = Math.Max(fontSize, _measurer.MeasureLineHeight(style));
            var box = new LayoutBox(width, height, height * 0.78D);
            if (!string.IsNullOrEmpty(text)) box.Commands.Add(LayoutCommand.TextCommand(text, 0D, 0D, width, height, fontSize));
            return box;
        }

        private LayoutBox Row(IReadOnlyList<OfficeMathExpression> children, double scale) {
            if (children.Count == 0) return Text(string.Empty, scale);
            var boxes = new List<LayoutBox>(children.Count);
            double baseline = 0D;
            double descent = 0D;
            double width = 0D;
            for (int index = 0; index < children.Count; index++) {
                LayoutBox child = Layout(children[index], scale);
                boxes.Add(child);
                baseline = Math.Max(baseline, child.Baseline);
                descent = Math.Max(descent, child.Height - child.Baseline);
                width += child.Width;
            }
            var result = new LayoutBox(width, baseline + descent, baseline);
            double x = 0D;
            for (int index = 0; index < boxes.Count; index++) {
                result.Add(boxes[index], x, baseline - boxes[index].Baseline);
                x += boxes[index].Width;
            }
            return result;
        }

        private LayoutBox Fraction(OfficeMathExpression expression, double scale) {
            LayoutBox numerator = Layout(expression.Children[0], scale * 0.92D);
            LayoutBox denominator = Layout(expression.Children[1], scale * 0.92D);
            double gap = _options.RuleGap * scale;
            double inset = Math.Max(2D * scale, _options.RuleThickness);
            double width = Math.Max(numerator.Width, denominator.Width) + inset * 2D;
            double ruleY = numerator.Height + gap;
            double denominatorY = ruleY + _options.RuleThickness + gap;
            var box = new LayoutBox(width, denominatorY + denominator.Height, denominatorY + denominator.Baseline);
            box.Add(numerator, (width - numerator.Width) / 2D, 0D);
            box.Commands.Add(LayoutCommand.Line(inset / 2D, ruleY, width - inset / 2D, ruleY));
            box.Add(denominator, (width - denominator.Width) / 2D, denominatorY);
            return box;
        }

        private LayoutBox Radical(OfficeMathExpression expression, double scale) {
            LayoutBox content = Layout(expression.Children[0], scale);
            LayoutBox radical = Text("√", scale * 1.1D);
            double indexWidth = 0D;
            LayoutBox? index = null;
            if (expression.Children.Count == 2) {
                index = Layout(expression.Children[1], scale * _options.ScriptScale);
                indexWidth = index.Width * 0.65D;
            }
            double gap = _options.RuleGap * scale;
            double top = gap + _options.RuleThickness;
            double width = indexWidth + radical.Width + content.Width + gap;
            double height = Math.Max(radical.Height, top + content.Height);
            double baseline = Math.Max(radical.Baseline, top + content.Baseline);
            var box = new LayoutBox(width, height, baseline);
            if (index != null) box.Add(index, 0D, 0D);
            box.Add(radical, indexWidth, baseline - radical.Baseline);
            double contentX = indexWidth + radical.Width;
            box.Add(content, contentX, top);
            box.Commands.Add(LayoutCommand.Line(contentX, gap, width, gap));
            return box;
        }

        private LayoutBox Scripts(OfficeMathExpression expression, double scale, bool hasSubscript, bool hasSuperscript) {
            LayoutBox basis = Layout(expression.Children[0], scale);
            int subIndex = hasSubscript ? 1 : -1;
            int superIndex = hasSuperscript ? (hasSubscript ? 2 : 1) : -1;
            LayoutBox? sub = subIndex >= 0 ? Layout(expression.Children[subIndex], scale * _options.ScriptScale) : null;
            LayoutBox? sup = superIndex >= 0 ? Layout(expression.Children[superIndex], scale * _options.ScriptScale) : null;
            double scriptWidth = Math.Max(sub?.Width ?? 0D, sup?.Width ?? 0D);
            double supHeight = sup?.Height ?? 0D;
            double baseline = Math.Max(basis.Baseline + supHeight * 0.45D, supHeight);
            double basisY = baseline - basis.Baseline;
            double subY = basisY + basis.Baseline + (basis.Height - basis.Baseline) * 0.35D;
            double height = Math.Max(basisY + basis.Height, sub == null ? 0D : subY + sub.Height);
            var box = new LayoutBox(basis.Width + scriptWidth, height, baseline);
            box.Add(basis, 0D, basisY);
            if (sup != null) box.Add(sup, basis.Width, 0D);
            if (sub != null) box.Add(sub, basis.Width, subY);
            return box;
        }

        private LayoutBox LeftScripts(OfficeMathExpression expression, double scale) {
            LayoutBox basis = Layout(expression.Children[0], scale);
            LayoutBox sub = Layout(expression.Children[1], scale * _options.ScriptScale);
            LayoutBox sup = Layout(expression.Children[2], scale * _options.ScriptScale);
            double scriptWidth = Math.Max(sub.Width, sup.Width);
            double baseline = Math.Max(basis.Baseline + sup.Height * 0.45D, sup.Height);
            double basisY = baseline - basis.Baseline;
            double subY = basisY + basis.Baseline + (basis.Height - basis.Baseline) * 0.35D;
            double height = Math.Max(basisY + basis.Height, subY + sub.Height);
            var box = new LayoutBox(scriptWidth + basis.Width, height, baseline);
            box.Add(sup, scriptWidth - sup.Width, 0D);
            box.Add(sub, scriptWidth - sub.Width, subY);
            box.Add(basis, scriptWidth, basisY);
            return box;
        }

        private LayoutBox Limit(OfficeMathExpression expression, double scale, bool over) {
            LayoutBox basis = Layout(expression.Children[0], scale);
            LayoutBox limit = Layout(expression.Children[1], scale * _options.ScriptScale);
            double gap = _options.RuleGap * scale;
            double width = Math.Max(basis.Width, limit.Width);
            double basisY = over ? limit.Height + gap : 0D;
            double limitY = over ? 0D : basis.Height + gap;
            var box = new LayoutBox(width, basis.Height + gap + limit.Height, basisY + basis.Baseline);
            box.Add(basis, (width - basis.Width) / 2D, basisY);
            box.Add(limit, (width - limit.Width) / 2D, limitY);
            return box;
        }

        private LayoutBox SlashedFraction(OfficeMathExpression expression, double scale) {
            LayoutBox numerator = Layout(expression.Children[0], scale * 0.92D);
            LayoutBox slash = Text("/", scale * 1.1D);
            LayoutBox denominator = Layout(expression.Children[1], scale * 0.92D);
            return CombineOnBaseline(CombineOnBaseline(numerator, slash, scale), denominator, scale);
        }

        private LayoutBox Nary(OfficeMathExpression expression, double scale) {
            LayoutBox symbol = Text(expression.Character ?? "∑", scale * 1.35D);
            LayoutBox? lower = expression.NaryLowerLimit != null ? Layout(expression.NaryLowerLimit, scale * _options.ScriptScale) : null;
            LayoutBox? upper = expression.NaryUpperLimit != null ? Layout(expression.NaryUpperLimit, scale * _options.ScriptScale) : null;
            double operatorWidth = Math.Max(symbol.Width, Math.Max(lower?.Width ?? 0D, upper?.Width ?? 0D));
            double upperHeight = upper?.Height ?? 0D;
            double operatorHeight = upperHeight + symbol.Height + (lower?.Height ?? 0D);
            var op = new LayoutBox(operatorWidth, operatorHeight, upperHeight + symbol.Baseline);
            if (upper != null) op.Add(upper, (operatorWidth - upper.Width) / 2D, 0D);
            op.Add(symbol, (operatorWidth - symbol.Width) / 2D, upperHeight);
            if (lower != null) op.Add(lower, (operatorWidth - lower.Width) / 2D, upperHeight + symbol.Height);
            LayoutBox content = Layout(expression.Children[0], scale);
            return CombineOnBaseline(op, content, 2D * scale);
        }

        private LayoutBox Delimited(OfficeMathExpression expression, double scale) {
            LayoutBox content = Layout(expression.Children[0], scale);
            double delimiterScale = Math.Max(scale, scale * content.Height / Math.Max(1D, Text("(", scale).Height));
            LayoutBox left = Text(expression.Character ?? "(", delimiterScale);
            LayoutBox right = Text(expression.SecondaryCharacter ?? ")", delimiterScale);
            return CombineOnBaseline(CombineOnBaseline(left, content, 1D * scale), right, 1D * scale);
        }

        private LayoutBox DelimiterList(OfficeMathExpression expression, double scale) {
            var content = new List<OfficeMathExpression>();
            for (int index = 0; index < expression.Children.Count; index++) {
                if (index > 0) content.Add(OfficeMath.Operator(expression.SeparatorCharacter ?? ","));
                content.Add(expression.Children[index]);
            }
            return Delimited(OfficeMath.Create(
                OfficeMathKind.Delimited,
                children: new[] { OfficeMath.Row(content.ToArray()) },
                character: expression.Character,
                secondaryCharacter: expression.SecondaryCharacter), scale);
        }

        private LayoutBox Stack(IReadOnlyList<OfficeMathExpression> rows, double scale) {
            var boxes = rows.Select(row => Layout(row, scale)).ToArray();
            double gap = _options.MatrixGap * scale;
            double width = boxes.Max(box => box.Width);
            double height = boxes.Sum(box => box.Height) + gap * Math.Max(0, boxes.Length - 1);
            double reference = height / 2D + Text("x", scale).Baseline / 2D;
            var result = new LayoutBox(width, height, reference);
            double y = 0D;
            for (int index = 0; index < boxes.Length; index++) {
                result.Add(boxes[index], (width - boxes[index].Width) / 2D, y);
                y += boxes[index].Height + gap;
            }
            return result;
        }

        private LayoutBox Matrix(OfficeMathExpression expression, double scale) {
            int rows = expression.RowCount;
            int columns = expression.ColumnCount;
            var cells = new LayoutBox[rows, columns];
            var widths = new double[columns];
            var heights = new double[rows];
            var baselines = new double[rows];
            var descents = new double[rows];
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    LayoutBox cell = Layout(expression.Children[row * columns + column], scale);
                    cells[row, column] = cell;
                    widths[column] = Math.Max(widths[column], cell.Width);
                    baselines[row] = Math.Max(baselines[row], cell.Baseline);
                    descents[row] = Math.Max(descents[row], cell.Height - cell.Baseline);
                }
                heights[row] = baselines[row] + descents[row];
            }
            double gap = _options.MatrixGap * scale;
            double tableWidth = Sum(widths) + gap * Math.Max(0, columns - 1);
            double tableHeight = Sum(heights) + gap * Math.Max(0, rows - 1);
            var table = new LayoutBox(tableWidth, tableHeight, tableHeight / 2D + Text("x", scale).Baseline / 2D);
            double y = 0D;
            for (int row = 0; row < rows; row++) {
                double x = 0D;
                for (int column = 0; column < columns; column++) {
                    LayoutBox cell = cells[row, column];
                    table.Add(cell, x + (widths[column] - cell.Width) / 2D, y + baselines[row] - cell.Baseline);
                    x += widths[column] + gap;
                }
                y += heights[row] + gap;
            }
            if (expression.Kind == OfficeMathKind.EquationArray) return table;
            LayoutBox left = Text("[", scale * Math.Max(1D, tableHeight / Math.Max(1D, Text("[", scale).Height)));
            LayoutBox right = Text("]", scale * Math.Max(1D, tableHeight / Math.Max(1D, Text("]", scale).Height)));
            return CombineOnBaseline(CombineOnBaseline(left, table, gap / 3D), right, gap / 3D);
        }

        private LayoutBox Accent(OfficeMathExpression expression, double scale) {
            LayoutBox content = Layout(expression.Children[0], scale);
            LayoutBox accent = Text(expression.Character ?? "^", scale * _options.ScriptScale);
            double width = Math.Max(content.Width, accent.Width);
            var box = new LayoutBox(width, accent.Height + content.Height, accent.Height + content.Baseline);
            box.Add(accent, (width - accent.Width) / 2D, 0D);
            box.Add(content, (width - content.Width) / 2D, accent.Height);
            return box;
        }

        private LayoutBox Bar(OfficeMathExpression contentExpression, double scale, bool over) {
            LayoutBox content = Layout(contentExpression, scale);
            double gap = _options.RuleGap * scale;
            double extra = gap + _options.RuleThickness;
            var box = new LayoutBox(content.Width, content.Height + extra, content.Baseline + (over ? extra : 0D));
            box.Add(content, 0D, over ? extra : 0D);
            double y = over ? _options.RuleThickness / 2D : content.Height + gap;
            box.Commands.Add(LayoutCommand.Line(0D, y, content.Width, y));
            return box;
        }

        private LayoutBox Box(OfficeMathExpression contentExpression, double scale) {
            LayoutBox content = Layout(contentExpression, scale);
            double inset = Math.Max(2D, _options.RuleGap) * scale;
            var box = new LayoutBox(content.Width + inset * 2D, content.Height + inset * 2D, content.Baseline + inset);
            box.Add(content, inset, inset);
            box.Commands.Add(LayoutCommand.Rectangle(0D, 0D, box.Width, box.Height));
            return box;
        }

        private static LayoutBox CombineOnBaseline(LayoutBox left, LayoutBox right, double gap) {
            double baseline = Math.Max(left.Baseline, right.Baseline);
            double height = baseline + Math.Max(left.Height - left.Baseline, right.Height - right.Baseline);
            var box = new LayoutBox(left.Width + gap + right.Width, height, baseline);
            box.Add(left, 0D, baseline - left.Baseline);
            box.Add(right, left.Width + gap, baseline - right.Baseline);
            return box;
        }

        private static double Sum(double[] values) { double total = 0D; for (int index = 0; index < values.Length; index++) total += values[index]; return total; }
    }

    private enum LayoutCommandKind { Text, Line, Rectangle }

    private sealed class LayoutCommand {
        internal LayoutCommandKind Kind { get; private set; }
        internal string? Text { get; private set; }
        internal double X { get; private set; }
        internal double Y { get; private set; }
        internal double X2 { get; private set; }
        internal double Y2 { get; private set; }
        internal double Width { get; private set; }
        internal double Height { get; private set; }
        internal double FontSize { get; private set; }

        internal static LayoutCommand TextCommand(string text, double x, double y, double width, double height, double fontSize) =>
            new LayoutCommand { Kind = LayoutCommandKind.Text, Text = text, X = x, Y = y, Width = width, Height = height, FontSize = fontSize };
        internal static LayoutCommand Line(double x1, double y1, double x2, double y2) =>
            new LayoutCommand { Kind = LayoutCommandKind.Line, X = x1, Y = y1, X2 = x2, Y2 = y2 };
        internal static LayoutCommand Rectangle(double x, double y, double width, double height) =>
            new LayoutCommand { Kind = LayoutCommandKind.Rectangle, X = x, Y = y, Width = width, Height = height };
        internal LayoutCommand Translate(double x, double y) => new LayoutCommand {
            Kind = Kind, Text = Text, X = X + x, Y = Y + y, X2 = X2 + x, Y2 = Y2 + y,
            Width = Width, Height = Height, FontSize = FontSize
        };
    }

    private sealed class LayoutBox {
        internal LayoutBox(double width, double height, double baseline) { Width = width; Height = height; Baseline = baseline; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Baseline { get; }
        internal List<LayoutCommand> Commands { get; } = new List<LayoutCommand>();
        internal void Add(LayoutBox child, double x, double y) {
            for (int index = 0; index < child.Commands.Count; index++) Commands.Add(child.Commands[index].Translate(x, y));
        }
    }
}
