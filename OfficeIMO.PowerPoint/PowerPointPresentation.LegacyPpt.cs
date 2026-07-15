using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private const double EmusPerLegacyMasterUnit = 1587.5d;

        /// <summary>Loads a binary `.ppt`, `.pot`, or `.pps` file into the normal editable PowerPoint model.</summary>
        public static PowerPointPresentation LoadLegacyPpt(string path, LegacyPptImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(path, options);
            return ProjectLoadedLegacyPpt(legacy, path,
                PowerPointPresentationLoadRouting.GetFormat(path, legacyDefault: true), new PowerPointLoadOptions());
        }

        /// <summary>Loads a binary PowerPoint stream into the normal editable PowerPoint model.</summary>
        public static PowerPointPresentation LoadLegacyPpt(Stream stream, LegacyPptImportOptions? options = null) {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(stream, options);
            return ProjectLoadedLegacyPpt(legacy, sourcePath: null, PowerPointFileFormat.Ppt,
                new PowerPointLoadOptions());
        }

        /// <summary>Loads a binary PowerPoint file and returns its projected presentation and import report.</summary>
        public static LegacyPptLoadResult LoadLegacyPptWithReport(string path, LegacyPptImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(path, options);
            try {
                return new LegacyPptLoadResult(ProjectLoadedLegacyPpt(legacy, path,
                    PowerPointPresentationLoadRouting.GetFormat(path, legacyDefault: true), new PowerPointLoadOptions()), legacy);
            } catch (InvalidDataException exception) {
                return new LegacyPptLoadResult(document: null, legacy, exception);
            }
        }

        /// <summary>Loads a binary PowerPoint stream and returns its projected presentation and import report.</summary>
        public static LegacyPptLoadResult LoadLegacyPptWithReport(Stream stream, LegacyPptImportOptions? options = null) {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(stream, options);
            try {
                return new LegacyPptLoadResult(ProjectLoadedLegacyPpt(legacy, sourcePath: null,
                    PowerPointFileFormat.Ppt, new PowerPointLoadOptions()), legacy);
            } catch (InvalidDataException exception) {
                return new LegacyPptLoadResult(document: null, legacy, exception);
            }
        }

        private static PowerPointPresentation LoadLegacyPptFromNormalFlow(byte[] bytes, string? sourcePath,
            Stream? sourceStream, PowerPointLoadOptions options) {
            if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && sourceStream == null
                && string.IsNullOrEmpty(sourcePath)) {
                throw new NotSupportedException("SaveOnDispose requires an associated destination for binary PowerPoint sources.");
            }
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes, options.LegacyPptImportOptions);
            PowerPointFileFormat sourceFormat = PowerPointPresentationLoadRouting.GetFormat(sourcePath, legacyDefault: true);
            return ProjectLoadedLegacyPpt(legacy, sourcePath, sourceFormat, options, sourceStream);
        }

        private static PowerPointPresentation ProjectLoadedLegacyPpt(LegacyPptPresentation legacy,
            string? sourcePath, PowerPointFileFormat sourceFormat, PowerPointLoadOptions loadOptions,
            Stream? sourceStream = null) {
            if (legacy == null) throw new ArgumentNullException(nameof(legacy));
            using PowerPointPresentation projected = Create();
            projected.SlideSize.SetSizeEmus(ToEmus(legacy.SlideWidth), ToEmus(legacy.SlideHeight));
            IReadOnlyDictionary<uint, LegacyPptLayoutTarget> layoutTargets =
                ProjectLegacyMasters(projected, legacy);

            foreach (LegacyPptSlide legacySlide in legacy.Slides) {
                PowerPointSlide slide = layoutTargets.TryGetValue(legacySlide.MasterId,
                    out LegacyPptLayoutTarget target)
                    ? projected.AddSlide(target.MasterIndex, target.LayoutIndex)
                    : projected.AddSlide();
                slide.Hidden = legacySlide.Hidden;
                foreach (LegacyPptShape shape in legacySlide.Shapes) {
                    ProjectLegacyShape(slide, shape);
                }
                if (!string.IsNullOrWhiteSpace(legacySlide.NotesText)) {
                    slide.Notes.Text = legacySlide.NotesText;
                }
            }

            byte[] packageBytes = projected.ToBytes();
            PowerPointPresentation presentation = LoadPackage(packageBytes, sourcePath, sourceStream, loadOptions);
            LegacyPptProjectionMap projectionMap = LegacyPptProjectionMap.Create(presentation, legacy);
            presentation.MarkLoadedFromLegacyPpt(sourcePath, legacy, projectionMap, sourceFormat);
            return presentation;
        }

        private static void ProjectLegacyShape(PowerPointSlide slide, LegacyPptShape shape) {
            long left = ToEmus(shape.Bounds.Left);
            long top = ToEmus(shape.Bounds.Top);
            long width = Math.Max(1L, ToEmus(shape.Bounds.Width));
            long height = Math.Max(1L, ToEmus(shape.Bounds.Height));
            switch (shape.Kind) {
                case LegacyPptShapeKind.TextBox:
                    PowerPointTextBox textBox = shape.PlaceholderKind == LegacyPptPlaceholderKind.Title
                        || shape.PlaceholderKind == LegacyPptPlaceholderKind.CenterTitle
                        || shape.PlaceholderKind == LegacyPptPlaceholderKind.VerticalTitle
                        ? slide.AddTitle(shape.Text, left, top, width, height)
                        : slide.AddTextBox(shape.Text, left, top, width, height);
                    PlaceholderValues? placeholder = MapPlaceholder(shape.PlaceholderKind);
                    if (placeholder.HasValue) textBox.PlaceholderType = placeholder.Value;
                    break;
                case LegacyPptShapeKind.Rectangle:
                    slide.AddShape(A.ShapeTypeValues.Rectangle, left, top, width, height);
                    break;
                case LegacyPptShapeKind.Ellipse:
                    slide.AddShape(A.ShapeTypeValues.Ellipse, left, top, width, height);
                    break;
                case LegacyPptShapeKind.Line:
                    slide.AddShape(A.ShapeTypeValues.Line, left, top, width, height);
                    break;
            }
        }

        private static PlaceholderValues? MapPlaceholder(LegacyPptPlaceholderKind placeholder) {
            switch (placeholder) {
                case LegacyPptPlaceholderKind.MasterTitle:
                case LegacyPptPlaceholderKind.Title: return PlaceholderValues.Title;
                case LegacyPptPlaceholderKind.MasterCenterTitle:
                case LegacyPptPlaceholderKind.CenterTitle: return PlaceholderValues.CenteredTitle;
                case LegacyPptPlaceholderKind.MasterSubtitle:
                case LegacyPptPlaceholderKind.Subtitle: return PlaceholderValues.SubTitle;
                case LegacyPptPlaceholderKind.MasterBody:
                case LegacyPptPlaceholderKind.Body: return PlaceholderValues.Body;
                case LegacyPptPlaceholderKind.VerticalTitle: return PlaceholderValues.Title;
                case LegacyPptPlaceholderKind.VerticalBody: return PlaceholderValues.Body;
                case LegacyPptPlaceholderKind.MasterNotesSlideImage:
                case LegacyPptPlaceholderKind.NotesSlideImage: return PlaceholderValues.SlideImage;
                case LegacyPptPlaceholderKind.MasterNotesBody:
                case LegacyPptPlaceholderKind.NotesBody: return PlaceholderValues.Body;
                case LegacyPptPlaceholderKind.MasterDate: return PlaceholderValues.DateAndTime;
                case LegacyPptPlaceholderKind.MasterSlideNumber: return PlaceholderValues.SlideNumber;
                case LegacyPptPlaceholderKind.MasterFooter: return PlaceholderValues.Footer;
                case LegacyPptPlaceholderKind.MasterHeader: return PlaceholderValues.Header;
                case LegacyPptPlaceholderKind.Graph: return PlaceholderValues.Chart;
                case LegacyPptPlaceholderKind.Table: return PlaceholderValues.Table;
                case LegacyPptPlaceholderKind.ClipArt: return PlaceholderValues.ClipArt;
                case LegacyPptPlaceholderKind.Media: return PlaceholderValues.Media;
                case LegacyPptPlaceholderKind.Picture: return PlaceholderValues.Picture;
                case LegacyPptPlaceholderKind.Object:
                case LegacyPptPlaceholderKind.OrganizationChart:
                case LegacyPptPlaceholderKind.VerticalObject: return PlaceholderValues.Object;
                default: return null;
            }
        }

        private static long ToEmus(int masterUnits) => checked((long)Math.Round(
            masterUnits * EmusPerLegacyMasterUnit, MidpointRounding.AwayFromZero));
    }
}
