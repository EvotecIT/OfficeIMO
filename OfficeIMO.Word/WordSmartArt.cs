using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a SmartArt diagram in a <see cref="WordDocument"/>.
    /// </summary>
    public class WordSmartArt : WordElement {
        private static int _docPrIdSeed = 1;

        private static UInt32Value GenerateDocPrId() {
            int id = Interlocked.Increment(ref _docPrIdSeed);
            return (UInt32Value)(uint)id;
        }

        internal Drawing _drawing;
        private readonly WordDocument _document;
        private readonly WordParagraph _paragraph;

        internal WordSmartArt(WordDocument document, WordParagraph paragraph, SmartArtType type) {
            _document = document;
            _paragraph = paragraph;

            InsertSmartArt(type);
        }

        internal WordSmartArt(WordDocument document, WordParagraph paragraph, Drawing drawing) {
            _document = document;
            _paragraph = paragraph;
            _drawing = drawing;
        }

        private void InsertSmartArt(SmartArtType type) {
            var mainPart = _document._wordprocessingDocument.MainDocumentPart!;

            var layoutPart = mainPart.AddNewPart<DiagramLayoutDefinitionPart>();
            layoutPart.LayoutDefinition = new LayoutDefinition();
            var colorsPart = mainPart.AddNewPart<DiagramColorsPart>();
            colorsPart.ColorsDefinition = new ColorsDefinition();
            var stylePart = mainPart.AddNewPart<DiagramStylePart>();
            stylePart.StyleDefinition = new StyleDefinition();
            var dataPart = mainPart.AddNewPart<DiagramDataPart>();
            dataPart.DataModelRoot = new DataModelRoot();

            var relLayout = mainPart.GetIdOfPart(layoutPart);
            var relColors = mainPart.GetIdOfPart(colorsPart);
            var relStyle = mainPart.GetIdOfPart(stylePart);
            var relData = mainPart.GetIdOfPart(dataPart);

            var graphic = new Graphic(new GraphicData(
                new RelationshipIds {
                    LayoutPart = relLayout,
                    StylePart = relStyle,
                    ColorPart = relColors,
                    DataPart = relData
                }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" });

            var inline = new Inline(
                new Extent { Cx = 5486400, Cy = 3200400 },
                new EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DocProperties { Id = GenerateDocPrId(), Name = "SmartArt" },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new GraphicFrameLocks { NoChangeAspect = true }),
                graphic) {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                };

            _drawing = new Drawing(inline);
            _paragraph.VerifyRun();
            _paragraph._run.Append(_drawing);
        }
    }
}
