using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// The chart members that still reach the Open XML SDK package: finding, adding and saving the chart part. Chart markup
    /// itself is read and written through <see cref="OfficeXmlNode"/>.
    /// </summary>
    public partial class WordChart {
        internal WordChart(WordDocument document, WordParagraph paragraph, WordDrawing drawing)
            : this(document, paragraph, OfficeXmlNode.FromSdk(drawing)) {
        }

        /// <summary>The chart part the drawing references, or null when it has none or the reference is broken.</summary>
        internal ChartPart? ChartPart {
            get {
                if (_drawing.SdkElement is not OpenXmlElement drawing) return null;
                string? id = ChartReferenceId(_drawing);
                if (string.IsNullOrWhiteSpace(id)) return null;
                try {
                    return WordPartOwnership.Resolve(_document, drawing).GetPartById(id!) as ChartPart;
                } catch (ArgumentOutOfRangeException) {
                    return null;
                } catch (InvalidOperationException) {
                    return null;
                } catch (OpenXmlPackageException) {
                    return null;
                }
            }
        }

        /// <summary>The chart's drawing as the converters still take it.</summary>
        internal WordDrawing? Drawing => _drawing.SdkElement as WordDrawing;

        /// <summary>The chart part as the shared chart reader and writer take it, or null.</summary>
        private IOfficeChartPart? SharedChartPart => ChartPart is { } part ? new OfficeSdkChartPart(part) : null;

        /// <summary>The chart part's <c>c:chartSpace</c>, or none.</summary>
        private OfficeXmlNode ChartSpaceNode => OfficeXmlNode.FromSdk(ChartPart?.ChartSpace);

        /// <summary>Writes the chart space to its part now, as edits through the SDK's loaded tree are otherwise held.</summary>
        private void SaveChartSpace() => ChartPart?.ChartSpace?.Save();

        /// <summary>The document theme's color scheme, which scheme colors in the chart resolve through, or none.</summary>
        private OfficeXmlNode ThemeColorScheme =>
            OfficeXmlNode.FromSdk(_document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme);

        /// <summary>The document theme's font scheme, or none.</summary>
        private OfficeXmlNode ThemeFontScheme =>
            OfficeXmlNode.FromSdk(_document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.FontScheme);

        /// <summary>Whether the document settings map theme color slots to other slots.</summary>
        private bool DocumentRemapsThemeColors() {
            var mapping = _document.MainDocumentPartRoot.DocumentSettingsPart?.Settings?
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.ColorSchemeMapping>();
            return mapping != null && (
                mapping.Background1?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Light1 ||
                mapping.Text1?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Dark1 ||
                mapping.Background2?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Light2 ||
                mapping.Text2?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Dark2 ||
                mapping.Accent1?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Accent1 ||
                mapping.Accent2?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Accent2 ||
                mapping.Accent3?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Accent3 ||
                mapping.Accent4?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Accent4 ||
                mapping.Accent5?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Accent5 ||
                mapping.Accent6?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Accent6 ||
                mapping.Hyperlink?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.Hyperlink ||
                mapping.FollowedHyperlink?.Value != DocumentFormat.OpenXml.Wordprocessing.ColorSchemeIndexValues.FollowedHyperlink);
        }

        /// <summary>Reads the drawing's package geometry.</summary>
        private bool TryReadDrawingLayout(out WordDrawingLayoutInfo snapshot) =>
            WordDrawingLayoutReader.TryRead(_drawing.SdkElement as WordDrawing, out snapshot);

        /// <summary>Sets the drawing's extent.</summary>
        private void SetDrawingExtent(long widthEmus, long? heightEmus) {
            if (_drawing.SdkElement is WordDrawing drawing) WordDrawingLayoutReader.SetExtent(drawing, widthEmus, heightEmus);
        }

        /// <summary>The paragraph's own element, for building and inserting runs.</summary>
        private static OfficeXmlNode ParagraphNode(WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._paragraph);

        /// <summary>
        /// Adds an empty chart part under the story part that holds <paramref name="paragraph"/> and returns its relationship
        /// id and new <c>c:chartSpace</c>.
        /// </summary>
        private (string RelationshipId, OfficeXmlNode ChartSpace) AddChartPart(WordParagraph paragraph) {
            OpenXmlPart owner = WordPartOwnership.Resolve(_document, paragraph._paragraph);
            ChartPart part = owner.AddNewPart<ChartPart>();
            SetNewChartSpace(part);
            return (owner.GetIdOfPart(part), OfficeXmlNode.FromSdk(part.ChartSpace));
        }

        /// <summary>Gives the chart part a new, empty <c>c:chartSpace</c> declaring the chart, drawing and relationship namespaces.</summary>
        private static void SetNewChartSpace(ChartPart part) {
            var space = new C.ChartSpace();
            space.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            space.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            space.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            part.ChartSpace = space;
        }

        /// <summary>Replaces the chart part's root with a new, empty one, as the shared chart writer fills a fresh root.</summary>
        private void ReplaceChartSpaceRoot() => SetNewChartSpace(ChartPart ?? throw new InvalidOperationException("The chart has no native drawing part."));

        /// <summary>Adds an embedded workbook package to the chart part and returns its relationship id.</summary>
        private string AddEmbeddedWorkbook() {
            ChartPart part = ChartPart ?? throw new InvalidOperationException("The chart has no native drawing part.");
            return part.GetIdOfPart(part.AddEmbeddedPackagePart(OfficeOpenXmlChartWriter.SharedWorkbookContentType));
        }

        /// <summary>Replaces the bytes of the embedded package <paramref name="relationshipId"/> names.</summary>
        private void WriteEmbeddedPackage(string relationshipId, byte[] data) {
            var package = (EmbeddedPackagePart)(ChartPart ?? throw new InvalidOperationException("The chart has no native drawing part."))
                .GetPartById(relationshipId);
            using var stream = new MemoryStream(data, writable: false);
            package.FeedData(stream);
        }

        /// <summary>
        /// Raises the shared axis id seed above every axis id in the document's charts, so charts added later don't reuse
        /// one. Axis ids in extension markup are not chart axes and are skipped.
        /// </summary>
        internal static void InitializeAxisIdSeed(WordprocessingDocument document) {
            uint max = (uint)System.Threading.Volatile.Read(ref _axisIdSeed);
            foreach (var part in document.MainDocumentPart?.ChartParts ?? Enumerable.Empty<ChartPart>()) {
                OfficeXmlNode chart = OfficeXmlNode.FromSdk(part.ChartSpace).Child(NsC, "chart");
                if (chart.IsNone) continue;
                OfficeXmlNode plot = chart.Child(NsC, "plotArea");
                if (plot.IsNone) continue;
                // Axis ids sit in the plot area's chart groups and axes. An id that isn't an unsigned number can't collide
                // with one this library allocates.
                for (OfficeXmlNode holder = plot.FirstChild; !holder.IsNone; holder = holder.NextSibling) {
                    if (holder.NamespaceId != NsC || !IsChartGroup(holder) && !IsAxis(holder)) continue;
                    foreach (OfficeXmlNode axis in holder.Children(NsC, "axId")) {
                        if (uint.TryParse(axis.Val(), System.Globalization.NumberStyles.AllowLeadingWhite | System.Globalization.NumberStyles.AllowTrailingWhite,
                                System.Globalization.CultureInfo.InvariantCulture, out uint value) && value > max) max = value;
                    }
                }
            }
            RaiseAxisIdSeed(max);
        }
    }
}
