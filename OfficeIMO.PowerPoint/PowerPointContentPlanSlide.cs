using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>One content slide rendered by the shared layout engine, including any required continuation slides.</summary>
    public sealed class PowerPointContentPlanSlide : PowerPointDeckPlanSlide {
        /// <summary>Creates a content slide. Auto selects its layout from the supplied blocks.</summary>
        public PowerPointContentPlanSlide(string? title, IEnumerable<PowerPointLayoutBlock> blocks,
            PowerPointLayout layout = PowerPointLayout.Auto, string? subtitle = null, string? seed = null)
            : base(title, subtitle, seed, allowEmptyTitle: true) {
            if (blocks == null) throw new ArgumentNullException(nameof(blocks));
            if (!Enum.IsDefined(typeof(PowerPointLayout), layout)) throw new ArgumentOutOfRangeException(nameof(layout));
            var items = blocks.ToList();
            if (items.Any(item => item == null)) throw new ArgumentException("Content blocks cannot be null.", nameof(blocks));
            var semantic = items.Where(item => !item.IsContentBlock).ToList();
            if (semantic.Count > 0) {
                bool singleKind = semantic.All(item => item is PowerPointLayoutKpi) ||
                    semantic.All(item => item is PowerPointLayoutCard) || semantic.All(item => item is PowerPointLayoutStep) ||
                    semantic.All(item => item is PowerPointLayoutLocation) || semantic.Count == 1 && semantic[0] is PowerPointLayoutQuote;
                bool caseStudy = semantic.All(item => item is PowerPointLayoutCard or PowerPointLayoutKpi) &&
                    semantic.OfType<PowerPointLayoutCard>().Count() <= PowerPointDeckPlanLimits.MaxCaseStudySections &&
                    semantic.OfType<PowerPointLayoutKpi>().Count() <= PowerPointDeckPlanLimits.MaxCaseStudyMetrics &&
                    semantic.OfType<PowerPointLayoutCard>().All(card => string.IsNullOrWhiteSpace(card.Image));
                if (semantic.Count != items.Count || !singleKind && !caseStudy)
                    throw new ArgumentException("Use separate slides for different semantic block kinds. A case study combines up to four narrative cards and three metrics.", nameof(blocks));
            }
            Blocks = items.AsReadOnly();
            Layout = layout;
        }

        /// <summary>Ordered content passed directly to the layout engine.</summary>
        public IReadOnlyList<PowerPointLayoutBlock> Blocks { get; }
        /// <summary>Requested layout, or Auto to select from content.</summary>
        public PowerPointLayout Layout { get; }
        /// <summary>Optional eyebrow above the title.</summary>
        public string? Eyebrow { get; set; }
        /// <summary>Label above the coverage panel.</summary>
        public string? MapLabel { get; set; }
        /// <summary>Metadata lines on an opening title slide.</summary>
        public IList<string> Metadata { get; } = new List<string>();
        /// <summary>Specific template layout for this slide; otherwise the composition's kind mapping applies.</summary>
        public PowerPointTemplateLayoutInfo? TemplateLayout { get; set; }
        /// <summary>Speaker notes; never added to the visible slide body.</summary>
        public string? Notes { get; set; }
        /// <summary>Takeaway below metrics, cards or process steps.</summary>
        public string? Takeaway { get; set; }
        /// <summary>Caption or source below the main content.</summary>
        public string? SourceNote { get; set; }
        /// <summary>Suppress footer text on this slide.</summary>
        public bool NoFooter { get; set; }
        /// <summary>Suppress the slide number on this slide.</summary>
        public bool NoSlideNumber { get; set; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Content;
        internal override int ContentItemCount => Blocks.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) => deck.AddContentSlide(this);

        internal PowerPointLayoutSlide ToLayout() {
            var slide = new PowerPointLayoutSlide {
                Title = Title, Subtitle = Subtitle, Layout = Layout, Eyebrow = Eyebrow,
                MapLabel = MapLabel, TemplateLayout = TemplateLayout,
                Notes = Notes, Takeaway = Takeaway, SourceNote = SourceNote,
                NoFooter = NoFooter, NoSlideNumber = NoSlideNumber
            };
            slide.Blocks.AddRange(Blocks);
            slide.Meta.AddRange(Metadata);
            return slide;
        }

        internal PowerPointLayout ResolvePreviewLayout(int position, bool isLast) =>
            PowerPointLayoutEngine.ResolveForPosition(ToLayout(), position == 0, isLast, position + 1);
    }

    public sealed partial class PowerPointDeckPlan {
        /// <summary>Adds content to the existing Compose workflow without a separate slide renderer.</summary>
        public PowerPointDeckPlan AddContent(string? title, IEnumerable<PowerPointLayoutBlock> blocks,
            PowerPointLayout layout = PowerPointLayout.Auto, string? subtitle = null) =>
            Add(new PowerPointContentPlanSlide(title, blocks, layout, subtitle));
    }
}
