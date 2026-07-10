using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>Executive-summary slide request.</summary>
    public sealed class PowerPointExecutiveSummaryPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointExecutiveSummarySlideOptions>? _configure;

        /// <summary>Creates an executive-summary slide request.</summary>
        public PowerPointExecutiveSummaryPlanSlide(string title, string? subtitle,
            PowerPointExecutiveSummaryContent content, string? seed = null,
            Action<PowerPointExecutiveSummarySlideOptions>? configure = null) : base(title, subtitle, seed) {
            Content = content ?? throw new ArgumentNullException(nameof(content));
            _configure = configure;
        }

        /// <summary>Summary content.</summary>
        public PowerPointExecutiveSummaryContent Content { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.ExecutiveSummary;
        internal override int ContentItemCount => Content.Metrics.Count + Content.Points.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddExecutiveSummarySlide(Title, Subtitle, Content, Seed, _configure);
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointExecutiveSummarySlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveExecutiveVariant(options, Content).ToString();
        }
    }

    /// <summary>Editable chart-story slide request.</summary>
    public sealed class PowerPointChartStoryPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointChartStorySlideOptions>? _configure;

        /// <summary>Creates a chart-story slide request.</summary>
        public PowerPointChartStoryPlanSlide(string title, string? subtitle, PowerPointChartStoryContent content,
            string? seed = null, Action<PowerPointChartStorySlideOptions>? configure = null)
            : base(title, subtitle, seed) {
            Content = content ?? throw new ArgumentNullException(nameof(content));
            _configure = configure;
        }

        /// <summary>Chart and narrative content.</summary>
        public PowerPointChartStoryContent Content { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.ChartStory;
        internal override int ContentItemCount => Content.SharedData.Series.Count + Content.Insights.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddChartStorySlide(Title, Subtitle, Content, Seed, _configure);
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointChartStorySlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveChartStoryVariant(options, Content).ToString();
        }
    }

    /// <summary>Semantic comparison slide request.</summary>
    public sealed class PowerPointComparisonPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointComparisonSlideOptions>? _configure;

        /// <summary>Creates a comparison slide request.</summary>
        public PowerPointComparisonPlanSlide(string title, string? subtitle,
            IEnumerable<PowerPointComparisonItem> items, string? seed = null,
            Action<PowerPointComparisonSlideOptions>? configure = null) : base(title, subtitle, seed) {
            Items = Materialize(items, nameof(items));
            _configure = configure;
        }

        /// <summary>Compared options.</summary>
        public IReadOnlyList<PowerPointComparisonItem> Items { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Comparison;
        internal override int ContentItemCount => Items.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddComparisonSlide(Title, Subtitle, Items, Seed, _configure);
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointComparisonSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveComparisonVariant(options, Items.Count).ToString();
        }
        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Items.Count < 2 || Items.Count > 4) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "Comparison.ItemCount", "Comparison slides require two to four options.");
            }
        }
    }

    /// <summary>Semantic screenshot-story slide request.</summary>
    public sealed class PowerPointScreenshotStoryPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointScreenshotStorySlideOptions>? _configure;

        /// <summary>Creates a screenshot-story slide request.</summary>
        public PowerPointScreenshotStoryPlanSlide(string title, string? subtitle, PowerPointImageAsset image,
            IEnumerable<string>? narrative = null, string? seed = null,
            Action<PowerPointScreenshotStorySlideOptions>? configure = null) : base(title, subtitle, seed) {
            Image = image ?? throw new ArgumentNullException(nameof(image));
            Narrative = (narrative ?? Array.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList().AsReadOnly();
            _configure = configure;
        }

        /// <summary>Semantic image asset.</summary>
        public PowerPointImageAsset Image { get; }
        /// <summary>Narrative points displayed with the image.</summary>
        public IReadOnlyList<string> Narrative { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.ScreenshotStory;
        internal override int ContentItemCount => Image.Annotations.Count + Narrative.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddScreenshotStorySlide(Title, Subtitle, Image, Narrative, Seed, _configure);
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointScreenshotStorySlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveScreenshotVariant(options, Image, Narrative).ToString();
        }
        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (!File.Exists(Image.Path)) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "ScreenshotStory.MissingImage", "Screenshot asset was not found: " + Image.Path);
            }
        }
    }

    /// <summary>Paginated editable appendix-table slide request.</summary>
    public sealed class PowerPointAppendixTablePlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointAppendixTableSlideOptions>? _configure;

        /// <summary>Creates an appendix-table slide request.</summary>
        public PowerPointAppendixTablePlanSlide(string title, string? subtitle, PowerPointTableData data,
            string? seed = null, Action<PowerPointAppendixTableSlideOptions>? configure = null)
            : base(title, subtitle, seed) {
            Data = data ?? throw new ArgumentNullException(nameof(data));
            _configure = configure;
        }

        /// <summary>Editable table data.</summary>
        public PowerPointTableData Data { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.AppendixTable;
        internal override int ContentItemCount => Data.Rows.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddAppendixTableSlide(Title, Subtitle, Data, Seed, _configure);
        internal override IEnumerable<PowerPointDeckPlanSlide> ExpandContinuations(
            PowerPointDeckContinuationOptions options) {
            IReadOnlyList<IReadOnlyList<IReadOnlyList<string>>> pages =
                PowerPointDeckContinuationOptions.Chunk(Data.Rows, options.AppendixRowsPerSlide);
            if (pages.Count <= 1) {
                yield return this;
                yield break;
            }
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
                var pageData = new PowerPointTableData(Data.Headers, pages[pageIndex]) {
                    Caption = Data.Caption,
                    Provenance = Data.Provenance,
                    Notes = Data.Notes
                };
                yield return new PowerPointAppendixTablePlanSlide(
                    options.CreateTitle(Title, pageIndex, pages.Count), Subtitle, pageData,
                    options.CreateSeed(Seed, Title, pageIndex), _configure);
            }
        }
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointAppendixTableSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveAppendixVariant(options, Data).ToString();
        }
        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Data.Rows.Count > PowerPointDeckPlanLimits.MaxAppendixTableRows) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "AppendixTable.TooManyRows", "Appendix tables support up to " +
                    PowerPointDeckPlanLimits.MaxAppendixTableRows + " rows per slide; use continuations.");
            }
        }
    }

    /// <summary>Editable architecture slide request.</summary>
    public sealed class PowerPointArchitecturePlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointArchitectureSlideOptions>? _configure;

        /// <summary>Creates an architecture slide request.</summary>
        public PowerPointArchitecturePlanSlide(string title, string? subtitle,
            PowerPointArchitectureContent content, string? seed = null,
            Action<PowerPointArchitectureSlideOptions>? configure = null) : base(title, subtitle, seed) {
            Content = content ?? throw new ArgumentNullException(nameof(content));
            _configure = configure;
        }

        /// <summary>Architecture graph content.</summary>
        public PowerPointArchitectureContent Content { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Architecture;
        internal override int ContentItemCount => Content.Nodes.Count + Content.Edges.Count;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddArchitectureSlide(Title, Subtitle, Content, Seed, _configure);
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointArchitectureSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveArchitectureVariant(options, Content).ToString();
        }
        internal override void Validate(int index, IList<PowerPointDeckPlanDiagnostic> diagnostics) {
            if (Content.Nodes.Count > 12) {
                AddDiagnostic(diagnostics, index, PowerPointDeckPlanDiagnosticSeverity.Error,
                    "Architecture.TooManyNodes", "Architecture slides support up to 12 nodes.");
            }
        }
    }

    /// <summary>Closing slide request.</summary>
    public sealed class PowerPointClosingPlanSlide : PowerPointDeckPlanSlide {
        private readonly Action<PowerPointClosingSlideOptions>? _configure;

        /// <summary>Creates a closing slide request.</summary>
        public PowerPointClosingPlanSlide(string title, PowerPointClosingContent content, string? seed = null,
            Action<PowerPointClosingSlideOptions>? configure = null) : base(title, null, seed) {
            Content = content ?? throw new ArgumentNullException(nameof(content));
            _configure = configure;
        }

        /// <summary>Closing content.</summary>
        public PowerPointClosingContent Content { get; }
        /// <inheritdoc />
        public override PowerPointDeckPlanSlideKind Kind => PowerPointDeckPlanSlideKind.Closing;
        internal override int ContentItemCount => 1;
        internal override PowerPointSlide AddTo(PowerPointDeckComposer deck) =>
            deck.AddClosingSlide(Title, Content, Seed, _configure);
        private protected override string ResolveLayoutVariant(PowerPointDeckDesign design, string slideSeed) {
            PowerPointClosingSlideOptions options = ConfigurePreview(design, slideSeed, _configure);
            return PowerPointDesignExtensions.ResolveClosingVariant(options, Content).ToString();
        }
    }
}
