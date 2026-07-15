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
                    master.Shapes.Count(shape => shape.Placeholder != null))
                + presentation.Slides.Sum(slide =>
                    slide.NotesPage?.Shapes.Count(shape => shape.Placeholder != null) ?? 0)
                + CountSpecialMasterPlaceholders(presentation.NotesMaster)
                + CountSpecialMasterPlaceholders(presentation.HandoutMaster);
            DistinctSlideLayoutCount = presentation.Slides
                .Select(slide => $"{slide.MasterId:X8}:{slide.LayoutType:X8}:"
                    + string.Join("-", slide.LayoutPlaceholderTypes
                        .Select(value => ((byte)value).ToString("X2"))))
                .Distinct(StringComparer.Ordinal)
                .Count();
            MasterTextStyleCount = presentation.Masters.Sum(master => master.TextMasterStyles.Count);
            MasterTextStyleLevelCount = presentation.Masters.Sum(master =>
                master.TextMasterStyles.Sum(style => style.Levels.Count));
            SpecialMasterCount = (presentation.NotesMaster == null ? 0 : 1)
                + (presentation.HandoutMaster == null ? 0 : 1);
            SpecialMasterShapeCount = CountSpecialMasterShapes(presentation.NotesMaster)
                + CountSpecialMasterShapes(presentation.HandoutMaster);
            BackgroundCount = presentation.Slides.Count(slide => slide.Background != null)
                + presentation.Slides.Count(slide => slide.NotesPage?.Background != null)
                + presentation.Masters.Count(master => master.Background != null)
                + CountSpecialMasterBackground(presentation.NotesMaster)
                + CountSpecialMasterBackground(presentation.HandoutMaster);
            ProjectableBackgroundCount = presentation.Slides.Count(slide =>
                    slide.Background?.HasProjectableFill == true)
                + presentation.Slides.Count(slide =>
                    slide.NotesPage?.Background?.HasProjectableFill == true)
                + presentation.Masters.Count(master =>
                    master.Background?.HasProjectableFill == true)
                + CountProjectableSpecialMasterBackground(presentation.NotesMaster)
                + CountProjectableSpecialMasterBackground(presentation.HandoutMaster);
            UnsupportedShapeCount = presentation.Slides.Sum(slide =>
                    slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.Unsupported)
                    + (slide.NotesPage?.Shapes.Count(shape =>
                        shape.Kind == LegacyPptShapeKind.Unsupported) ?? 0));
            NotesSlideCount = presentation.Slides.Count(slide => slide.NotesPage != null);
            NotesPageShapeCount = presentation.Slides.Sum(slide =>
                slide.NotesPage?.Shapes.Count ?? 0);
            HeaderFooterScopeCount = (presentation.SlideHeaderFooterDefaults == null ? 0 : 1)
                + (presentation.NotesHeaderFooterDefaults == null ? 0 : 1)
                + presentation.Masters.Count(master => master.HeaderFooter != null)
                + presentation.Slides.Count(slide => slide.HeaderFooter != null);
            TransitionCount = presentation.Slides.Count(slide =>
                slide.Transition != null);
            TransitionSoundCount = presentation.Slides.Count(slide =>
                slide.Transition?.PlaySound == true
                || slide.Transition?.StopSound == true);
            CommentCount = presentation.Slides.Sum(slide => slide.Comments.Count);
            CommentAuthorCount = presentation.Slides.SelectMany(slide => slide.Comments)
                .Select(comment => (comment.Author, comment.Initials))
                .Distinct()
                .Count();
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

        /// <summary>Gets the number of decoded notes and handout masters.</summary>
        public int SpecialMasterCount { get; }

        /// <summary>Gets the number of decoded notes- and handout-master shapes.</summary>
        public int SpecialMasterShapeCount { get; }

        /// <summary>Gets the number of decoded OfficeArt background shapes.</summary>
        public int BackgroundCount { get; }

        /// <summary>Gets the number of background shapes with a projectable primary fill.</summary>
        public int ProjectableBackgroundCount { get; }

        /// <summary>Gets the preserve-only unsupported shape count.</summary>
        public int UnsupportedShapeCount { get; }

        /// <summary>Gets the number of slides with imported speaker notes.</summary>
        public int NotesSlideCount { get; }

        /// <summary>Gets the number of decoded notes-page drawing shapes.</summary>
        public int NotesPageShapeCount { get; }

        /// <summary>Gets the number of decoded document, master, and per-slide header/footer scopes.</summary>
        public int HeaderFooterScopeCount { get; }

        /// <summary>Gets the number of decoded slide-show transition atoms.</summary>
        public int TransitionCount { get; }

        /// <summary>Gets the number of transitions that play or stop sound.</summary>
        public int TransitionSoundCount { get; }

        /// <summary>Gets the number of decoded legacy review comments.</summary>
        public int CommentCount { get; }

        /// <summary>Gets the number of distinct embedded comment authors.</summary>
        public int CommentAuthorCount { get; }

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

        private static int CountSpecialMasterShapes(LegacyPptSpecialMaster? master) =>
            master?.Shapes.Count ?? 0;

        private static int CountSpecialMasterPlaceholders(LegacyPptSpecialMaster? master) =>
            master?.Shapes.Count(shape => shape.Placeholder != null) ?? 0;

        private static int CountSpecialMasterBackground(LegacyPptSpecialMaster? master) =>
            master?.Background == null ? 0 : 1;

        private static int CountProjectableSpecialMasterBackground(
            LegacyPptSpecialMaster? master) =>
            master?.Background?.HasProjectableFill == true ? 1 : 0;
    }
}
