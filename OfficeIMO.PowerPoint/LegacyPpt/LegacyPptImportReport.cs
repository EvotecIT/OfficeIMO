using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>Provides a compact inventory of a binary PowerPoint import.</summary>
    public sealed class LegacyPptImportReport {
        internal LegacyPptImportReport(LegacyPptPresentation presentation) {
            SlideCount = presentation.Slides.Count;
            ShapeCount = presentation.Slides.Sum(slide => slide.Shapes.Count);
            TextShapeCount = presentation.Slides.Sum(slide =>
                slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.TextBox));
            PictureShapeCount = presentation.Slides.Sum(slide =>
                slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.Picture));
            BlipStoreEntryCount = presentation.BlipStoreEntries.Count;
            ImportableBlipCount = presentation.BlipStoreEntries.Count(entry => entry.HasImportableImage);
            FontCount = presentation.Fonts.Count;
            EmbeddedFontCount = presentation.Fonts.Count(font => font.HasEmbeddedData);
            TextRulerCount = presentation.Slides.Sum(slide =>
                slide.Shapes.Count(shape => shape.TextBody.HasRulerRecord));
            PlaceholderShapeCount = presentation.Slides.Sum(slide =>
                slide.Shapes.Count(shape => shape.Placeholder != null))
                + presentation.Masters.Sum(master =>
                    master.Shapes.Count(shape => shape.Placeholder != null));
            DistinctSlideLayoutCount = presentation.Slides
                .Select(slide => $"{slide.MasterId:X8}:{slide.LayoutType:X8}:"
                    + string.Join("-", slide.LayoutPlaceholderTypes
                        .Select(value => ((byte)value).ToString("X2"))))
                .Distinct(StringComparer.Ordinal)
                .Count();
            MasterTextStyleCount = presentation.Masters.Sum(master => master.TextMasterStyles.Count);
            MasterTextStyleLevelCount = presentation.Masters.Sum(master =>
                master.TextMasterStyles.Sum(style => style.Levels.Count));
            UnsupportedShapeCount = presentation.Slides.Sum(slide =>
                slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.Unsupported));
            NotesSlideCount = presentation.Slides.Count(slide => !string.IsNullOrWhiteSpace(slide.NotesText));
            WarningCount = presentation.Diagnostics.Count(diagnostic =>
                diagnostic.Severity == LegacyPptDiagnosticSeverity.Warning);
            ErrorCount = presentation.Diagnostics.Count(diagnostic =>
                diagnostic.Severity == LegacyPptDiagnosticSeverity.Error);
            UserEditCount = presentation.Package.UserEdits.Count;
            PersistObjectCount = presentation.Package.PersistObjects.Count;
            CompoundStreamCount = presentation.Package.CompoundStreamCount;
        }

        /// <summary>Gets the presentation slide count.</summary>
        public int SlideCount { get; }

        /// <summary>Gets the decoded shape count.</summary>
        public int ShapeCount { get; }

        /// <summary>Gets the decoded text-shape count.</summary>
        public int TextShapeCount { get; }

        /// <summary>Gets the number of slide picture frames with importable image data.</summary>
        public int PictureShapeCount { get; }

        /// <summary>Gets the number of document-level OfficeArt BLIP store entries.</summary>
        public int BlipStoreEntryCount { get; }

        /// <summary>Gets the number of BLIP entries that can be projected as Open XML images.</summary>
        public int ImportableBlipCount { get; }

        /// <summary>Gets the number of decoded document font entries.</summary>
        public int FontCount { get; }

        /// <summary>Gets the number of font entries with preserved embedded font data.</summary>
        public int EmbeddedFontCount { get; }

        /// <summary>Gets the number of slide text shapes that contain a TextRulerAtom.</summary>
        public int TextRulerCount { get; }

        /// <summary>Gets the number of decoded placeholder shapes across slides and masters.</summary>
        public int PlaceholderShapeCount { get; }

        /// <summary>Gets the number of distinct master/layout-signature combinations used by slides.</summary>
        public int DistinctSlideLayoutCount { get; }

        /// <summary>Gets the number of decoded base master text styles.</summary>
        public int MasterTextStyleCount { get; }

        /// <summary>Gets the number of decoded base master text-style levels.</summary>
        public int MasterTextStyleLevelCount { get; }

        /// <summary>Gets the preserve-only unsupported shape count.</summary>
        public int UnsupportedShapeCount { get; }

        /// <summary>Gets the number of slides with imported speaker notes.</summary>
        public int NotesSlideCount { get; }

        /// <summary>Gets the warning count.</summary>
        public int WarningCount { get; }

        /// <summary>Gets the error count.</summary>
        public int ErrorCount { get; }

        /// <summary>Gets the number of UserEditAtom revisions retained from the source.</summary>
        public int UserEditCount { get; }

        /// <summary>Gets the number of live persist objects retained from the source.</summary>
        public int PersistObjectCount { get; }

        /// <summary>Gets the number of exact compound streams retained from the source.</summary>
        public int CompoundStreamCount { get; }

        /// <summary>Gets whether projection to PPTX has known conversion loss.</summary>
        public bool HasConversionLoss => WarningCount > 0 || UnsupportedShapeCount > 0;
    }
}
