using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>The designed slide layouts (Build/Visual/powerpoint-layouts.json).</summary>
    internal enum PowerPointLayoutKind {
        Auto,
        Title,
        Section,
        TitleAndContent,
        TwoContent,
        Comparison,
        Kpi,
        Cards,
        Process,
        Quote,
        ImageFull,
        Closing,
        CaseStudy,
        Logos,
        Coverage
    }

    /// <summary>
    ///     Content of one slide for the layout system. It mirrors the PowerShell <c>PowerPointSlide</c> command: a title, optional subtitle and
    ///     eyebrow, and an ordered list of content blocks. Geometry is never part of the input.
    /// </summary>
    internal sealed class PowerPointLayoutSlide {
        internal string? Title { get; set; }
        internal string? Subtitle { get; set; }
        internal string? Eyebrow { get; set; }
        internal string? MapLabel { get; set; }
        internal PowerPointLayoutKind Layout { get; set; }
        internal List<PowerPointLayoutBlock> Blocks { get; } = new();
        internal string? Notes { get; set; }
        internal bool NoFooter { get; set; }
        internal bool NoSlideNumber { get; set; }

        /// <summary>Caption or source line under the content (caption style).</summary>
        internal string? SourceNote { get; set; }

        /// <summary>A specific template layout to use instead of the mapped built-in layout.</summary>
        internal PowerPointTemplateLayoutInfo? TemplateLayout { get; set; }

        /// <summary>Lines in the Title layout's meta panel (author, date).</summary>
        internal List<string> Meta { get; } = new();

        /// <summary>Takeaway under KPI tiles or process steps.</summary>
        internal string? Takeaway { get; set; }

        /// <summary>Overrides the deck-wide position test (first/last slide) used by automatic selection.</summary>
        internal bool? IsFirst { get; set; }

        internal bool? IsLast { get; set; }

        // Continuation state, set by the overflow pass.
        internal string? ContinuedSuffix { get; set; }
        internal int ListStart { get; set; } = 1;
        internal (int First, int Last, int Total)? TableRange { get; set; }
        internal double? ListSize { get; set; }
        internal double? TitleSize { get; set; }
        internal int StepStart { get; set; } = 1;
        internal int CardStart { get; set; } = 1;

        internal IEnumerable<T> Items<T>() where T : PowerPointLayoutBlock => Blocks.OfType<T>();

        internal PowerPointLayoutSlide CloneHeader() {
            var copy = new PowerPointLayoutSlide {
                Title = Title,
                Subtitle = Subtitle,
                Eyebrow = Eyebrow,
                MapLabel = MapLabel,
                Layout = Layout,
                NoFooter = NoFooter,
                NoSlideNumber = NoSlideNumber,
                IsFirst = false,
                IsLast = IsLast,
                TitleSize = TitleSize,
                ListSize = ListSize,
                TemplateLayout = TemplateLayout,
                SourceNote = SourceNote
            };
            copy.Meta.AddRange(Meta);
            return copy;
        }
    }

    internal abstract class PowerPointLayoutBlock {
        /// <summary>True for blocks counted by selection (lists, text, tables, charts, images, columns).</summary>
        internal virtual bool IsContentBlock => true;
    }

    internal sealed class PowerPointLayoutListItem {
        internal PowerPointLayoutListItem(string text, int level = 1) {
            Text = text ?? string.Empty;
            Level = Math.Max(1, Math.Min(2, level));
        }

        internal string Text { get; }
        internal int Level { get; }
    }

    internal sealed class PowerPointLayoutList : PowerPointLayoutBlock {
        internal PowerPointLayoutList(IEnumerable<PowerPointLayoutListItem> items, bool numbered = false) {
            Items = items.ToList();
            Numbered = numbered;
        }

        internal PowerPointLayoutList(IEnumerable<string> items, bool numbered = false)
            : this(items.Where(i => !string.IsNullOrWhiteSpace(i)).Select(i => new PowerPointLayoutListItem(i)), numbered) {
        }

        internal List<PowerPointLayoutListItem> Items { get; }
        internal bool Numbered { get; }
    }

    internal sealed class PowerPointLayoutText : PowerPointLayoutBlock {
        internal PowerPointLayoutText(IEnumerable<string> paragraphs) {
            Paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        }

        internal PowerPointLayoutText(string paragraph) : this(new[] { paragraph }) {
        }

        internal List<string> Paragraphs { get; }
        internal int Length => Paragraphs.Sum(p => p.Length);
    }

    internal sealed class PowerPointLayoutTable : PowerPointLayoutBlock {
        internal PowerPointLayoutTable(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string>> rows) {
            Headers = headers;
            Rows = rows;
        }

        internal IReadOnlyList<string> Headers { get; }
        internal IReadOnlyList<IReadOnlyList<string>> Rows { get; }
        internal int? MaxRows { get; set; }
        internal bool NoSplit { get; set; }
        internal string? Caption { get; set; }

        internal PowerPointLayoutTable WithRows(IReadOnlyList<IReadOnlyList<string>> rows) =>
            new(Headers, rows) { MaxRows = MaxRows, NoSplit = NoSplit, Caption = Caption };
    }

    internal sealed class PowerPointLayoutChart : PowerPointLayoutBlock {
        internal PowerPointLayoutChart(OfficeChartKind kind, OfficeChartData data) {
            Kind = kind;
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        internal OfficeChartKind Kind { get; }
        internal OfficeChartData Data { get; }
        internal string? Title { get; set; }
        internal string? AlternativeText { get; set; }
    }

    internal sealed class PowerPointLayoutImage : PowerPointLayoutBlock {
        internal PowerPointLayoutImage(string path, string? alternativeText) {
            Path = path;
            AlternativeText = alternativeText;
        }

        internal string Path { get; }
        internal string? AlternativeText { get; }
        internal double FocalX { get; set; } = 0.5D;
        internal double FocalY { get; set; } = 0.5D;
        internal bool Contain { get; set; }
        internal string? Caption { get; set; }

        /// <summary>Numbered markers at fractional positions of the source image.</summary>
        internal List<(double X, double Y, string Label)> Annotations { get; } = new();
    }

    internal sealed class PowerPointLayoutColumn : PowerPointLayoutBlock {
        internal PowerPointLayoutColumn(string? title, IEnumerable<PowerPointLayoutBlock> blocks) {
            Title = title;
            Blocks = blocks.ToList();
        }

        internal string? Title { get; }
        internal List<PowerPointLayoutBlock> Blocks { get; }
    }

    internal sealed class PowerPointLayoutQuote : PowerPointLayoutBlock {
        internal PowerPointLayoutQuote(string text, string? attribution) {
            Text = text;
            Attribution = attribution;
        }

        internal string Text { get; }
        internal string? Attribution { get; }
        internal override bool IsContentBlock => false;
    }

    internal sealed class PowerPointLayoutKpi : PowerPointLayoutBlock {
        internal PowerPointLayoutKpi(string value, string label, string? change = null, string? color = null) {
            Value = value;
            Label = label;
            Change = change;
            Color = color;
        }

        internal string Value { get; }
        internal string Label { get; }
        internal string? Change { get; }

        /// <summary>Hex colour or role name (Success, Warning, Danger, Accent1…).</summary>
        internal string? Color { get; }

        internal override bool IsContentBlock => false;
    }

    internal sealed class PowerPointLayoutCard : PowerPointLayoutBlock {
        internal PowerPointLayoutCard(string title, string? text = null, IEnumerable<string>? items = null,
            string? image = null, string? color = null) {
            Title = title;
            Text = text;
            Items = (items ?? Array.Empty<string>()).Where(i => !string.IsNullOrWhiteSpace(i)).ToList();
            Image = image;
            Color = color;
        }

        internal string Title { get; }
        internal string? Text { get; }
        internal List<string> Items { get; }
        internal string? Image { get; }
        internal string? Color { get; }
        internal bool IsLogo => !string.IsNullOrWhiteSpace(Image) && string.IsNullOrWhiteSpace(Text) && Items.Count == 0;
        internal override bool IsContentBlock => false;
    }

    internal sealed class PowerPointLayoutStep : PowerPointLayoutBlock {
        internal PowerPointLayoutStep(string title, string? text = null, string? number = null) {
            Title = title;
            Text = text;
            Number = number;
        }

        internal string Title { get; }
        internal string? Text { get; }
        internal string? Number { get; }
        internal override bool IsContentBlock => false;
    }

    internal sealed class PowerPointLayoutLocation : PowerPointLayoutBlock {
        internal PowerPointLayoutLocation(string name, double left, double top, string? detail = null) {
            Name = name;
            Left = Math.Max(0D, Math.Min(1D, left));
            Top = Math.Max(0D, Math.Min(1D, top));
            Detail = detail;
        }

        internal string Name { get; }
        internal double Left { get; }
        internal double Top { get; }
        internal string? Detail { get; }
        internal override bool IsContentBlock => false;
    }

    /// <summary>A free-form drawing hosted inside a designed slide's content zone (architecture diagrams, custom composition).</summary>
    internal sealed class PowerPointLayoutDrawing : PowerPointLayoutBlock {
        internal PowerPointLayoutDrawing(Action<PowerPointSlide, PowerPointLayoutContext, PowerPointLayoutRect> draw) {
            Draw = draw;
        }

        internal Action<PowerPointSlide, PowerPointLayoutContext, PowerPointLayoutRect> Draw { get; }
    }
}
