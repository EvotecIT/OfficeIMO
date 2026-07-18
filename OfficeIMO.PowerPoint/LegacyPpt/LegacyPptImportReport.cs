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
            PictureBulletCount = presentation.PictureBullets.Count;
            FontCount = presentation.Fonts.Count;
            EmbeddedFontCount = presentation.Fonts.Count(font => font.HasEmbeddedData);
            SoundCount = presentation.Sounds.Count;
            ImportableSoundCount = presentation.Sounds.Count(sound =>
                sound.HasData && sound.ContentType != null);
            EmbeddedOleObjectCount = presentation.OleObjects.Count;
            EmbeddedOleObjectByteCount = presentation.OleObjects.Sum(
                ole => ole.Length);
            CompressedEmbeddedOleObjectCount = presentation.OleObjects.Count(
                ole => ole.WasCompressed);
            LinkedOleObjectCount = presentation.LinkedOleObjects.Count;
            LinkedOleObjectByteCount = presentation.LinkedOleObjects.Sum(
                ole => ole.Length);
            CompressedLinkedOleObjectCount = presentation.LinkedOleObjects
                .Count(ole => ole.WasCompressed);
            ActiveXControlCount = presentation.ActiveXControls.Count;
            ActiveXControlByteCount = presentation.ActiveXControls.Sum(
                control => control.Length);
            CompressedActiveXControlCount = presentation.ActiveXControls
                .Count(control => control.WasCompressed);
            MediaObjectCount = presentation.Media.Count;
            EmbeddedWaveMediaCount = presentation.Media.Count(media =>
                media.Kind == LegacyPptMediaKind.EmbeddedWaveAudio);
            ProjectableMediaCount = presentation.Media.Count(media =>
                media.HasProjectableAudio);
            LinkedOrDeviceMediaCount = presentation.Media.Count(media =>
                media.Kind != LegacyPptMediaKind.EmbeddedWaveAudio);
            VbaProjectCount = presentation.VbaProject == null ? 0 : 1;
            VbaProjectByteCount = presentation.VbaProject?.Length ?? 0;
            VbaProjectWasCompressed = presentation.VbaProject?.WasCompressed
                == true;
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
            RoundTripThemeCount = presentation.Masters.Count(master =>
                    master.RoundTripTheme != null)
                + presentation.Slides.Count(slide =>
                    slide.RoundTripTheme != null)
                + presentation.Slides.Count(slide =>
                    slide.NotesPage?.RoundTripTheme != null)
                + CountSpecialMasterTheme(presentation.NotesMaster)
                + CountSpecialMasterTheme(presentation.HandoutMaster);
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
            HyperlinkTargetCount = presentation.Hyperlinks.Count;
            HyperlinkScreenTipCount = presentation.Hyperlinks.Count(hyperlink =>
                hyperlink.ScreenTip != null);
            HyperlinkExtensionFlagCount = presentation.Hyperlinks.Count(hyperlink =>
                hyperlink.ExtensionFlags != 0);
            CustomShowCount = presentation.CustomShows.Count;
            CustomShowSlideEntryCount = presentation.CustomShows.Sum(show =>
                show.SlideIds.Count);
            LegacyPptShape[] interactiveShapes = EnumerateShapes(presentation).ToArray();
            AnimationCount = interactiveShapes.Count(shape => shape.Animation != null);
            AnimationSoundCount = interactiveShapes.Count(shape =>
                shape.Animation?.PlaysSound == true || shape.Animation?.StopsSound == true);
            ShapeInteractionCount = interactiveShapes.Sum(shape => shape.Interactions.Count);
            TextInteractionCount = interactiveShapes.Sum(shape =>
                shape.TextBody.Interactions.Count);
            TextFieldCount = interactiveShapes.Sum(shape =>
                shape.TextBody.Fields.Count);
            WarningCount = presentation.Diagnostics.Count(diagnostic =>
                diagnostic.Severity == LegacyPptDiagnosticSeverity.Warning);
            ErrorCount = presentation.Diagnostics.Count(diagnostic =>
                diagnostic.Severity == LegacyPptDiagnosticSeverity.Error);
            UserEditCount = presentation.Package.UserEdits.Count;
            PersistObjectCount = presentation.Package.PersistObjects.Count;
            CompoundStreamCount = presentation.Package.CompoundStreamCount;
            WasEncryptedSource = presentation.WasEncryptedSource;
            EncryptionKeySizeBits = presentation.EncryptionKeySizeBits;
            EncryptedDocumentProperties =
                presentation.EncryptedDocumentProperties;
        }

        /// <summary>Gets the presentation slide count.</summary>
        public int SlideCount { get; }

        /// <summary>Gets whether the imported package was protected with RC4 CryptoAPI password encryption.</summary>
        public bool WasEncryptedSource { get; }

        /// <summary>Gets the imported RC4 key size in bits, when encrypted.</summary>
        public int? EncryptionKeySizeBits { get; }

        /// <summary>
        /// Gets whether an encrypted source protected its document-property streams,
        /// or <see langword="null"/> for an unencrypted source.
        /// </summary>
        public bool? EncryptedDocumentProperties { get; }

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

        /// <summary>Gets the number of decoded PPT9 picture-bullet images.</summary>
        public int PictureBulletCount { get; }

        /// <summary>Gets the number of decoded document font entries.</summary>
        public int FontCount { get; }

        /// <summary>Gets the number of font entries with preserved embedded font data.</summary>
        public int EmbeddedFontCount { get; }

        /// <summary>Gets the number of decoded document-level sound entries.</summary>
        public int SoundCount { get; }

        /// <summary>Gets the number of sounds that can be projected as Open XML audio.</summary>
        public int ImportableSoundCount { get; }

        /// <summary>Gets the number of decoded embedded OLE objects.</summary>
        public int EmbeddedOleObjectCount { get; }

        /// <summary>Gets the decoded compound-storage byte total.</summary>
        public int EmbeddedOleObjectByteCount { get; }

        /// <summary>Gets the number of compressed embedded OLE persist records.</summary>
        public int CompressedEmbeddedOleObjectCount { get; }

        /// <summary>Gets the number of decoded linked OLE objects.</summary>
        public int LinkedOleObjectCount { get; }

        /// <summary>Gets the decoded linked-cache compound-storage byte total.</summary>
        public int LinkedOleObjectByteCount { get; }

        /// <summary>Gets the number of compressed linked OLE cache persist records.</summary>
        public int CompressedLinkedOleObjectCount { get; }

        /// <summary>Gets the number of decoded ActiveX controls.</summary>
        public int ActiveXControlCount { get; }

        /// <summary>Gets the decoded ActiveX compound-storage byte total.</summary>
        public int ActiveXControlByteCount { get; }

        /// <summary>Gets the number of compressed ActiveX control persist records.</summary>
        public int CompressedActiveXControlCount { get; }

        /// <summary>Gets the number of decoded audio and movie objects.</summary>
        public int MediaObjectCount { get; }

        /// <summary>Gets the number of embedded WAV media definitions.</summary>
        public int EmbeddedWaveMediaCount { get; }

        /// <summary>Gets the number of embedded media objects that project editably.</summary>
        public int ProjectableMediaCount { get; }

        /// <summary>Gets the number of linked, path-based, or device-based media objects.</summary>
        public int LinkedOrDeviceMediaCount { get; }

        /// <summary>Gets the number of decoded presentation VBA projects.</summary>
        public int VbaProjectCount { get; }

        /// <summary>Gets the decompressed byte length of the decoded VBA project.</summary>
        public int VbaProjectByteCount { get; }

        /// <summary>Gets whether the decoded VBA persist object used the compressed storage form.</summary>
        public bool VbaProjectWasCompressed { get; }

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

        /// <summary>Gets the number of decoded DrawingML round-trip theme scopes.</summary>
        public int RoundTripThemeCount { get; }

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

        /// <summary>Gets the number of decoded document-level hyperlink targets.</summary>
        public int HyperlinkTargetCount { get; }

        /// <summary>Gets the number of decoded hyperlinks with PowerPoint 2000+ screen tips.</summary>
        public int HyperlinkScreenTipCount { get; }

        /// <summary>Gets the number of decoded hyperlinks with nonzero PowerPoint 2000+ flags.</summary>
        public int HyperlinkExtensionFlagCount { get; }

        /// <summary>Gets the number of decoded named custom shows.</summary>
        public int CustomShowCount { get; }

        /// <summary>Gets the number of ordered slide entries across named custom shows.</summary>
        public int CustomShowSlideEntryCount { get; }

        /// <summary>Gets the number of decoded classic shape and text animations.</summary>
        public int AnimationCount { get; }

        /// <summary>Gets the number of classic animations that play or stop sound.</summary>
        public int AnimationSoundCount { get; }

        /// <summary>Gets the number of decoded shape-level click and mouse-over actions.</summary>
        public int ShapeInteractionCount { get; }

        /// <summary>Gets the number of decoded text-range interactions.</summary>
        public int TextInteractionCount { get; }

        /// <summary>Gets the number of decoded dynamic text metacharacters.</summary>
        public int TextFieldCount { get; }

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

        private static int CountSpecialMasterTheme(LegacyPptSpecialMaster? master) =>
            master?.RoundTripTheme == null ? 0 : 1;

        private static int CountProjectableSpecialMasterBackground(
            LegacyPptSpecialMaster? master) =>
            master?.Background?.HasProjectableFill == true ? 1 : 0;

        private static IEnumerable<LegacyPptShape> EnumerateShapes(
            LegacyPptPresentation presentation) {
            IEnumerable<LegacyPptShape> roots = presentation.Slides
                .SelectMany(slide => slide.Shapes)
                .Concat(presentation.Slides.SelectMany(slide =>
                    slide.NotesPage?.Shapes ?? Array.Empty<LegacyPptShape>()))
                .Concat(presentation.Masters.SelectMany(master => master.Shapes))
                .Concat(presentation.NotesMaster?.Shapes ?? Array.Empty<LegacyPptShape>())
                .Concat(presentation.HandoutMaster?.Shapes ?? Array.Empty<LegacyPptShape>());
            foreach (LegacyPptShape shape in roots) {
                foreach (LegacyPptShape item in EnumerateShape(shape)) yield return item;
            }
        }

        private static IEnumerable<LegacyPptShape> EnumerateShape(LegacyPptShape shape) {
            yield return shape;
            foreach (LegacyPptShape child in shape.Children) {
                foreach (LegacyPptShape item in EnumerateShape(child)) yield return item;
            }
        }
    }
}
