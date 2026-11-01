using System;
using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>Stable diagnostic codes raised by the slide layout system.</summary>
    internal static class PowerPointLayoutDiagnosticCodes {
        internal const string LayoutContentMismatch = "layout-content-mismatch";
        internal const string TitleLong = "title-long";
        internal const string TitleTooLong = "title-too-long";
        internal const string TableOverflow = "table-overflow";
        internal const string TableTooWide = "table-too-wide";
        internal const string SlideContentSplit = "slide-content-split";
        internal const string ListLevelFlattened = "list-level-flattened";
        internal const string TemplateLayoutMissing = "template-layout-missing";
        internal const string TextSplitParagraph = "text-split-paragraph";
        internal const string LocationsTruncated = "coverage-list-truncated";
    }

    /// <summary>
    ///     Deck-level state of the layout system: the deck title used by footers, the current section (eyebrow
    ///     of following content slides), section numbering and the diagnostics raised while composing.
    /// </summary>
    internal sealed class PowerPointLayoutDeck {
        private readonly HashSet<string> _onceCodes = new(StringComparer.Ordinal);

        internal PowerPointLayoutDeck(PowerPointDocument document) {
            Document = document;
        }

        internal PowerPointDocument Document { get; }

        /// <summary>Deck title shown in the footer; the first Title slide's title unless set.</summary>
        internal string? DeckTitle { get; set; }

        /// <summary>Explicit footer text; null means the deck title, empty turns the footer text off.</summary>
        internal string? FooterText { get; set; }

        internal bool ShowSlideNumber { get; set; } = true;

        /// <summary>Suppress footer text throughout a composed deck.</summary>
        internal bool NoFooter { get; set; }

        /// <summary>Default deterministic variation seed; a content slide may override it.</summary>
        internal string? Seed { get; set; }

        /// <summary>Author and date lines used when a title slide supplies no explicit metadata.</summary>
        internal List<string> TitleMetadata { get; } = new();

        internal string? CurrentSection { get; set; }

        internal int SectionCount { get; set; }

        /// <summary>Continuation wording for the title, with {0} the title, {1} the part and {2} the count.</summary>
        internal string? ContinuationFormat { get; set; }

        internal List<OfficeDiagnostic> Diagnostics { get; } = new();

        /// <summary>Zero-based index of the slide being composed when each diagnostic was raised.</summary>
        internal List<int> DiagnosticSlides { get; } = new();

        internal string? EffectiveFooterText => FooterText ?? DeckTitle;

        /// <summary>Captures composition state so a failed logical slide cannot affect subsequent layouts.</summary>
        internal Action CaptureRollback() {
            string? title = DeckTitle, section = CurrentSection;
            int sections = SectionCount, diagnosticCount = Diagnostics.Count;
            var onceCodes = new HashSet<string>(_onceCodes, StringComparer.Ordinal);
            Action rollbackSections = Document.CaptureSectionRollback();
            return () => {
                rollbackSections();
                DeckTitle = title;
                CurrentSection = section;
                SectionCount = sections;
                Diagnostics.RemoveRange(diagnosticCount, Diagnostics.Count - diagnosticCount);
                DiagnosticSlides.RemoveRange(diagnosticCount, DiagnosticSlides.Count - diagnosticCount);
                _onceCodes.Clear();
                _onceCodes.UnionWith(onceCodes);
            };
        }

        internal void Report(string code, OfficeDiagnosticSeverity severity, string message, string? location = null,
            bool oncePerDeck = false) {
            if (oncePerDeck && !_onceCodes.Add(code)) return;
            Diagnostics.Add(new OfficeDiagnostic(code, severity, message, location, "PowerPoint"));
            DiagnosticSlides.Add(Document.Slides.Count);
        }
    }
}
