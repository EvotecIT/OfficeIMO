using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>The designed slide layouts (Build/Visual/powerpoint-layouts.json).</summary>
    public enum PowerPointLayout {
        /// <summary>Choose a layout from content and deck position.</summary>
        Auto,
        /// <summary>Opening title and metadata.</summary>
        Title,
        /// <summary>Section divider.</summary>
        Section,
        /// <summary>Title with one content zone.</summary>
        TitleAndContent,
        /// <summary>Two content zones.</summary>
        TwoContent,
        /// <summary>Two titled comparison columns.</summary>
        Comparison,
        /// <summary>Metric tiles and takeaway.</summary>
        Kpi,
        /// <summary>Content card grid.</summary>
        Cards,
        /// <summary>Process steps.</summary>
        Process,
        /// <summary>Quotation and attribution.</summary>
        Quote,
        /// <summary>Full-slide image.</summary>
        ImageFull,
        /// <summary>Closing statement or actions.</summary>
        Closing,
        /// <summary>Narrative cards and result metrics.</summary>
        CaseStudy,
        /// <summary>Logo or proof wall.</summary>
        Logos,
        /// <summary>Location coverage panel.</summary>
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
        internal PowerPointLayout Layout { get; set; }
        internal List<PowerPointLayoutBlock> Blocks { get; } = new();
        /// <summary>Native figures already occupy this slide; an empty flowing body is not an empty slide.</summary>
        internal bool HasFixedContent { get; set; }
        internal string? Notes { get; set; }
        internal bool NoFooter { get; set; }
        internal bool NoSlideNumber { get; set; }
        /// <summary>Explicit metadata beside the slide number for this slide and its continuations.</summary>
        internal string? FooterRightText { get; set; }

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

        /// <summary>Logical plan position used for layout selection; physical continuations do not change previews.</summary>
        internal int? SelectionDeckSize { get; set; }

        // Continuation state, set by the overflow pass.
        internal string? ContinuedSuffix { get; set; }
        internal int ListStart { get; set; } = 1;
        internal (int First, int Last, int Total)? TableRange { get; set; }
        internal double? ListSize { get; set; }
        internal double? TitleSize { get; set; }
        internal int StepStart { get; set; } = 1;
        internal int CardStart { get; set; } = 1;
        internal double? CardBodySize { get; set; }

        internal IEnumerable<T> Items<T>() where T : PowerPointLayoutBlock => Blocks.OfType<T>();

        internal PowerPointLayoutSlide CloneHeader() {
            var copy = new PowerPointLayoutSlide {
                Title = Title,
                Subtitle = Subtitle,
                Eyebrow = Eyebrow,
                MapLabel = MapLabel,
                Layout = Layout,
                HasFixedContent = HasFixedContent,
                NoFooter = NoFooter,
                NoSlideNumber = NoSlideNumber,
                FooterRightText = FooterRightText,
                IsFirst = false,
                IsLast = IsLast,
                SelectionDeckSize = SelectionDeckSize,
                TitleSize = TitleSize,
                ListSize = ListSize,
                TemplateLayout = TemplateLayout,
                SourceNote = SourceNote
            };
            copy.Meta.AddRange(Meta);
            return copy;
        }
    }

    /// <summary>A free-form drawing hosted inside a designed slide's content zone (architecture diagrams, custom composition).</summary>
    internal sealed class PowerPointLayoutDrawing : PowerPointLayoutBlock {
        internal PowerPointLayoutDrawing(Action<PowerPointSlide, PowerPointLayoutContext, PowerPointLayoutRect> draw) {
            Draw = draw;
        }

        internal Action<PowerPointSlide, PowerPointLayoutContext, PowerPointLayoutRect> Draw { get; }
    }
}
