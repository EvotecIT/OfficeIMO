using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>One semantic block consumed directly by the shared slide layout engine.</summary>
    public abstract class PowerPointLayoutBlock {
        private protected PowerPointLayoutBlock() { }
        internal static OfficeColor? ParseColor(string? value) => string.IsNullOrWhiteSpace(value) ? null : OfficeColor.Parse(value!);
        /// <summary>True for blocks counted by selection (lists, text, tables, charts, images, columns).</summary>
        internal virtual bool IsContentBlock => true;
    }

    /// <summary>One list item and its nesting level.</summary>
    public sealed class PowerPointLayoutListItem {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutListItem(string text, int level = 1) {
            Text = text ?? string.Empty;
            Level = Math.Max(1, Math.Min(2, level));
        }

        /// <summary>Text to render.</summary>
        public string Text { get; }
        /// <summary>One-based nesting level, clamped to one or two.</summary>
        public int Level { get; }
    }

    /// <summary>A numbered or bulleted list.</summary>
    public sealed class PowerPointLayoutList : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutList(IEnumerable<PowerPointLayoutListItem> items, bool numbered = false) {
            Items = (items ?? throw new ArgumentNullException(nameof(items))).ToList().AsReadOnly();
            Numbered = numbered;
        }

        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutList(IEnumerable<string> items, bool numbered = false)
            : this((items ?? throw new ArgumentNullException(nameof(items))).Where(i => !string.IsNullOrWhiteSpace(i)).Select(i => new PowerPointLayoutListItem(i)), numbered) {
        }

        /// <summary>Items in display order.</summary>
        public IReadOnlyList<PowerPointLayoutListItem> Items { get; }
        /// <summary>Whether to use numbers instead of bullets.</summary>
        public bool Numbered { get; }
    }

    /// <summary>One or more body paragraphs.</summary>
    public sealed class PowerPointLayoutText : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutText(IEnumerable<string> paragraphs) {
            Paragraphs = (paragraphs ?? throw new ArgumentNullException(nameof(paragraphs))).Where(p => !string.IsNullOrWhiteSpace(p)).ToList().AsReadOnly();
        }

        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutText(string paragraph) : this(new[] { paragraph }) {
        }

        /// <summary>Body paragraphs in display order.</summary>
        public IReadOnlyList<string> Paragraphs { get; }
        internal int Length => Paragraphs.Sum(p => p.Length);
    }

    /// <summary>A table with repeated headings and automatic continuation.</summary>
    public sealed class PowerPointLayoutTable : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutTable(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
            Headers = (headers ?? throw new ArgumentNullException(nameof(headers))).ToArray();
            Rows = (rows ?? throw new ArgumentNullException(nameof(rows))).Select(row => (IReadOnlyList<string>)(row ?? throw new ArgumentException("Table rows cannot be null.", nameof(rows))).ToArray()).ToArray();
            if (Headers.Count == 0) throw new ArgumentException("At least one table header is required.", nameof(headers));
            if (Rows.Any(row => row.Count != Headers.Count))
                throw new ArgumentException("Every table row must match the header count.", nameof(rows));
        }

        /// <summary>Column headings.</summary>
        public IReadOnlyList<string> Headers { get; }
        /// <summary>Table data rows.</summary>
        public IReadOnlyList<IReadOnlyList<string>> Rows { get; }
        private int? _maximumRows;
        /// <summary>Optional maximum data rows per physical slide.</summary>
        public int? MaximumRows {
            get => _maximumRows;
            set => _maximumRows = value.HasValue && value.Value < 1
                ? throw new ArgumentOutOfRangeException(nameof(value), "MaximumRows must be positive when supplied.")
                : value;
        }
        /// <summary>Keep the table on one slide and report any overflow.</summary>
        public bool NoSplit { get; set; }
        /// <summary>Optional caption below the content.</summary>
        public string? Caption { get; set; }

        internal PowerPointLayoutTable WithRows(IReadOnlyList<IReadOnlyList<string>> rows) =>
            new(Headers, rows) { MaximumRows = MaximumRows, NoSplit = NoSplit, Caption = Caption };
    }

    /// <summary>An editable chart backed by shared chart data.</summary>
    public sealed class PowerPointLayoutChart : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutChart(OfficeChartKind kind, OfficeChartData data) {
            Kind = kind;
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        /// <summary>Shared chart kind.</summary>
        public OfficeChartKind Kind { get; }
        /// <summary>Data for the editable chart.</summary>
        public OfficeChartData Data { get; }
        /// <summary>Content title.</summary>
        public string? Title { get; set; }
        /// <summary>Accessible description of the visual.</summary>
        public string? AlternativeText { get; set; }
    }

    /// <summary>An image fitted to its content zone.</summary>
    public sealed class PowerPointLayoutImage : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutImage(string path, string? alternativeText = null) {
            Path = path ?? throw new ArgumentNullException(nameof(path));
            AlternativeText = alternativeText;
        }

        /// <summary>Local image path.</summary>
        public string Path { get; }
        /// <summary>Accessible description of the visual.</summary>
        public string? AlternativeText { get; }
        /// <summary>Horizontal focal point as a fraction of image width.</summary>
        public double FocalX { get; set; } = 0.5D;
        /// <summary>Vertical focal point as a fraction of image height.</summary>
        public double FocalY { get; set; } = 0.5D;
        /// <summary>Fit the whole image rather than crop to fill.</summary>
        public bool Contain { get; set; }
        /// <summary>Optional caption below the content.</summary>
        public string? Caption { get; set; }

        /// <summary>Numbered markers at fractional positions of the source image.</summary>
        internal List<(double X, double Y, string Label)> Annotations { get; } = new();
    }

    /// <summary>A column containing ordered content blocks.</summary>
    public sealed class PowerPointLayoutColumn : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutColumn(string? title, IEnumerable<PowerPointLayoutBlock> blocks) {
            Title = title;
            Blocks = (blocks ?? throw new ArgumentNullException(nameof(blocks))).ToList().AsReadOnly();
            if (Blocks.Any(block => block == null || !block.IsContentBlock || block is PowerPointLayoutColumn) ||
                PowerPointLayoutEngine.SplitLede(Blocks).Blocks.Count > 1)
                throw new ArgumentException("A column supports one body block, optionally preceded by a single lede paragraph. Use separate slides for additional content.", nameof(blocks));
        }

        /// <summary>Content title.</summary>
        public string? Title { get; }
        /// <summary>Content blocks in display order.</summary>
        public IReadOnlyList<PowerPointLayoutBlock> Blocks { get; }
    }

    /// <summary>A quotation and attribution.</summary>
    public sealed class PowerPointLayoutQuote : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutQuote(string text, string? attribution = null) {
            Text = text ?? throw new ArgumentNullException(nameof(text));
            Attribution = attribution;
        }

        /// <summary>Text to render.</summary>
        public string Text { get; }
        /// <summary>Quotation attribution.</summary>
        public string? Attribution { get; }
        internal override bool IsContentBlock => false;
    }

    /// <summary>A formatted metric, label and optional change.</summary>
    public sealed class PowerPointLayoutKpi : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutKpi(string value, string label, string? change = null, OfficeColor? color = null) {
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Label = label ?? throw new ArgumentNullException(nameof(label));
            Change = change;
            Color = color;
        }

        /// <summary>Formatted metric value.</summary>
        public string Value { get; }
        /// <summary>Metric label.</summary>
        public string Label { get; }
        /// <summary>Optional formatted change.</summary>
        public string? Change { get; }

        /// <summary>Shared colour or theme role (Success, Warning, Danger, Accent1…).</summary>
        public OfficeColor? Color { get; }

        internal override bool IsContentBlock => false;
    }

    /// <summary>A narrative card or logo tile.</summary>
    public sealed class PowerPointLayoutCard : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutCard(string title, string? text = null, IEnumerable<string>? items = null,
            string? image = null, OfficeColor? color = null) {
            Title = title;
            Text = text;
            Items = (items ?? Array.Empty<string>()).Where(i => !string.IsNullOrWhiteSpace(i)).ToList().AsReadOnly();
            Image = image;
            Color = color;
        }

        /// <summary>Content title.</summary>
        public string Title { get; }
        /// <summary>Text to render.</summary>
        public string? Text { get; }
        /// <summary>Items in display order.</summary>
        public IReadOnlyList<string> Items { get; }
        /// <summary>Optional local image path.</summary>
        public string? Image { get; }
        /// <summary>Optional shared colour or theme role.</summary>
        public OfficeColor? Color { get; }
        internal bool IsLogo => !string.IsNullOrWhiteSpace(Image) && string.IsNullOrWhiteSpace(Text) && Items.Count == 0;
        internal override bool IsContentBlock => false;
    }

    /// <summary>A numbered process step.</summary>
    public sealed class PowerPointLayoutStep : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutStep(string title, string? text = null, string? number = null) {
            Title = title;
            Text = text;
            Number = number;
        }

        /// <summary>Content title.</summary>
        public string Title { get; }
        /// <summary>Text to render.</summary>
        public string? Text { get; }
        /// <summary>Optional display number.</summary>
        public string? Number { get; }
        internal override bool IsContentBlock => false;
    }

    /// <summary>A named location at fractional coordinates on a coverage panel.</summary>
    public sealed class PowerPointLayoutLocation : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutLocation(string name, double left, double top, string? detail = null) {
            Name = name;
            Left = Math.Max(0D, Math.Min(1D, left));
            Top = Math.Max(0D, Math.Min(1D, top));
            Detail = detail;
        }

        /// <summary>Location name.</summary>
        public string Name { get; }
        /// <summary>Horizontal fraction from zero to one.</summary>
        public double Left { get; }
        /// <summary>Vertical fraction from zero to one.</summary>
        public double Top { get; }
        /// <summary>Optional location detail.</summary>
        public string? Detail { get; }
        internal override bool IsContentBlock => false;
    }

}
