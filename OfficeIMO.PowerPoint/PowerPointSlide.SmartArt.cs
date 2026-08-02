using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        internal const string SmartArtColorsDefinitionXml = """
            <dgm:colorsDef uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <dgm:title val="" />
              <dgm:desc val="" />
              <dgm:catLst><dgm:cat type="accent1" pri="11200" /></dgm:catLst>
              <dgm:styleLbl name="node0"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1" /></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="lt1" /></dgm:linClrLst><dgm:effectClrLst /><dgm:txLinClrLst /><dgm:txFillClrLst /><dgm:txEffectClrLst /></dgm:styleLbl>
              <dgm:styleLbl name="lnNode1"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1" /></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="lt1" /></dgm:linClrLst><dgm:effectClrLst /><dgm:txLinClrLst /><dgm:txFillClrLst /><dgm:txEffectClrLst /></dgm:styleLbl>
              <dgm:styleLbl name="alignNode1"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1" /></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="accent1" /></dgm:linClrLst><dgm:effectClrLst /><dgm:txLinClrLst /><dgm:txFillClrLst /><dgm:txEffectClrLst /></dgm:styleLbl>
            </dgm:colorsDef>
            """;

        internal const string SmartArtStyleDefinitionXml = """
            <dgm:styleDef uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <dgm:title val="" />
              <dgm:desc val="" />
              <dgm:catLst><dgm:cat type="simple" pri="10100" /></dgm:catLst>
              <dgm:styleLbl name="node">
                <dgm:style>
                  <a:lnRef idx="2"><a:schemeClr val="accent1" /></a:lnRef>
                  <a:fillRef idx="1"><a:schemeClr val="accent1" /></a:fillRef>
                  <a:effectRef idx="0"><a:schemeClr val="accent1" /></a:effectRef>
                  <a:fontRef idx="minor"><a:schemeClr val="lt1" /></a:fontRef>
                </dgm:style>
              </dgm:styleLbl>
            </dgm:styleDef>
            """;

        /// <summary>
        ///     Adds a native SmartArt diagram to the slide.
        /// </summary>
        public PowerPointSmartArt AddSmartArt(PowerPointSmartArtType type = PowerPointSmartArtType.BasicProcess,
            long left = 0L, long top = 0L, long width = 5486400L, long height = 3200400L) {
            return AddSmartArt(type, new[] { "[Text]" }, left, top, width, height);
        }

        /// <summary>
        ///     Adds a native SmartArt diagram populated with editable semantic node text.
        /// </summary>
        public PowerPointSmartArt AddSmartArt(PowerPointSmartArtType type, IEnumerable<string> nodeTexts,
            long left = 0L, long top = 0L, long width = 5486400L, long height = 3200400L) {
            if (width <= 0 || width > MaximumDrawingCoordinate) {
                throw new ArgumentOutOfRangeException(nameof(width),
                    "Width must be representable by DrawingML.");
            }
            if (height <= 0 || height > MaximumDrawingCoordinate) {
                throw new ArgumentOutOfRangeException(nameof(height),
                    "Height must be representable by DrawingML.");
            }

            List<string> nodes = NormalizeSmartArtNodes(nodeTexts);
            uint shapeId = AllocateShapeId();
            var (layoutRelId, colorsRelId, styleRelId, dataRelId) = AddSmartArtParts(
                type, nodes, width, height);
            string name = GenerateUniqueName("SmartArt");
            GraphicFrame frame = CreateSmartArtFrame(shapeId, name, layoutRelId, colorsRelId, styleRelId,
                dataRelId, left, top, width, height);

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.Append(frame);

            return TrackShape(new PowerPointSmartArt(frame, _slidePart));
        }

        /// <summary>
        ///     Adds a native SmartArt diagram to the slide using a layout box.
        /// </summary>
        public PowerPointSmartArt AddSmartArt(PowerPointSmartArtType type, PowerPointLayoutBox layout) {
            return AddSmartArt(type, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>Adds populated native SmartArt using a layout box.</summary>
        public PowerPointSmartArt AddSmartArt(PowerPointSmartArtType type, IEnumerable<string> nodeTexts,
            PowerPointLayoutBox layout) {
            return AddSmartArt(type, nodeTexts, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        private (string layoutRelId, string colorsRelId, string styleRelId, string dataRelId) AddSmartArtParts(
            PowerPointSmartArtType type, IReadOnlyList<string> nodeTexts,
            long width, long height) {
            switch (type) {
                case PowerPointSmartArtType.BasicProcess:
                case PowerPointSmartArtType.BasicHierarchy:
                case PowerPointSmartArtType.BasicCycle:
                case PowerPointSmartArtType.BasicList:
                case PowerPointSmartArtType.BasicMatrix:
                case PowerPointSmartArtType.BasicPyramid:
                case PowerPointSmartArtType.BasicRelationship:
                    return AddSemanticSmartArtParts(type, nodeTexts,
                        width, height);
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type,
                        "Unsupported SmartArt type.");
            }
        }

        private (string layoutRelId, string colorsRelId, string styleRelId, string dataRelId)
            AddSemanticSmartArtParts(PowerPointSmartArtType type,
                IReadOnlyList<string> nodeTexts, long width, long height) {
            DiagramLayoutDefinitionPart layoutPart = _slidePart.AddNewPart<DiagramLayoutDefinitionPart>();
            PopulateSmartArtLayout(layoutPart, type, nodeTexts.Count,
                width / (double)height);
            DiagramColorsPart colorsPart = _slidePart.AddNewPart<DiagramColorsPart>();
            PopulateSmartArtColors(colorsPart);
            DiagramStylePart stylePart = _slidePart.AddNewPart<DiagramStylePart>();
            PopulateSmartArtStyle(stylePart);
            DiagramDataPart dataPart = _slidePart.AddNewPart<DiagramDataPart>();
            PopulateSmartArtData(dataPart, type, nodeTexts);

            return (
                _slidePart.GetIdOfPart(layoutPart)!,
                _slidePart.GetIdOfPart(colorsPart)!,
                _slidePart.GetIdOfPart(stylePart)!,
                _slidePart.GetIdOfPart(dataPart)!);
        }

        private static GraphicFrame CreateSmartArtFrame(uint id, string name, string layoutRelId, string colorsRelId,
            string styleRelId, string dataRelId, long left, long top, long width, long height) {
            return new GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = id, Name = name },
                    new NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()),
                new Transform(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(
                    new A.GraphicData(
                        new Dgm.RelationshipIds {
                            LayoutPart = layoutRelId,
                            StylePart = styleRelId,
                            ColorPart = colorsRelId,
                            DataPart = dataRelId
                        }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" }));
        }

        private static void PopulateSmartArtLayout(DiagramLayoutDefinitionPart part,
            PowerPointSmartArtType type, int nodeCount, double aspectRatio) {
            part.LayoutDefinition = CreateSmartArtLayoutDefinition(type,
                nodeCount, aspectRatio);
        }

        private static void PopulateSmartArtData(DiagramDataPart part, PowerPointSmartArtType type,
            IReadOnlyList<string> nodeTexts) {
            string docId = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
            var pointXml = new StringBuilder();
            var connectionXml = new StringBuilder();
            var childIds = new List<string>(nodeTexts.Count);
            for (int index = 0; index < nodeTexts.Count; index++) {
                childIds.Add("{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}");
            }
            bool usesRootedTopology = type == PowerPointSmartArtType.BasicHierarchy
                || type == PowerPointSmartArtType.BasicRelationship;
            for (int index = 0; index < nodeTexts.Count; index++) {
                string childId = childIds[index];
                string connectionId = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
                pointXml.Append("<dgm:pt modelId=\"").Append(childId).Append("\">")
                    .Append("<dgm:prSet phldr=\"0\" />")
                    .Append("<dgm:spPr /><dgm:t><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t>")
                    .Append(SecurityElement.Escape(nodeTexts[index]))
                    .Append("</a:t></a:r><a:endParaRPr lang=\"en-US\" /></a:p></dgm:t></dgm:pt>");
                string sourceId = usesRootedTopology && index > 0 ? childIds[0] : docId;
                int sourceOrder = usesRootedTopology && index > 0 ? index - 1 : index;
                connectionXml.Append("<dgm:cxn modelId=\"").Append(connectionId)
                    .Append("\" srcId=\"").Append(sourceId).Append("\" destId=\"").Append(childId)
                    .Append("\" srcOrd=\"").Append(sourceOrder).Append("\" destOrd=\"0\" />");
            }

            string xml = $"""
                <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <dgm:ptLst>
                    <dgm:pt modelId="{docId}" type="doc">
                      <dgm:prSet loTypeId="{GetSmartArtLayoutId(type)}" loCatId="{GetSmartArtCategory(type)}" qsTypeId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" qsCatId="simple" csTypeId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2" csCatId="accent1" phldr="0" />
                      <dgm:spPr />
                      <dgm:t><a:bodyPr /><a:lstStyle /><a:p><a:endParaRPr lang="en-US" /></a:p></dgm:t>
                    </dgm:pt>
                    {pointXml}
                  </dgm:ptLst>
                  <dgm:cxnLst>
                    {connectionXml}
                  </dgm:cxnLst>
                  <dgm:bg />
                  <dgm:whole />
                </dgm:dataModel>
                """;
            using MemoryStream stream = new(Encoding.UTF8.GetBytes(xml));
            part.FeedData(stream);
        }

        private static List<string> NormalizeSmartArtNodes(IEnumerable<string> nodeTexts) {
            if (nodeTexts == null) throw new ArgumentNullException(nameof(nodeTexts));
            List<string> nodes = nodeTexts.Select(text => (text ?? string.Empty).Trim())
                .Where(text => text.Length > 0).ToList();
            if (nodes.Count == 0) throw new ArgumentException("At least one SmartArt node is required.", nameof(nodeTexts));
            if (nodes.Count > 32) throw new ArgumentException("SmartArt workflows support at most 32 nodes.", nameof(nodeTexts));
            foreach (string node in nodes) {
                PowerPointXmlValueValidator.ValidateCharacters(node,
                    nameof(nodeTexts), "SmartArt node text");
            }
            return nodes;
        }

        private static string GetSmartArtLayoutId(PowerPointSmartArtType type) {
            switch (type) {
                case PowerPointSmartArtType.BasicHierarchy:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1";
                case PowerPointSmartArtType.BasicCycle:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/cycle2";
                case PowerPointSmartArtType.BasicList:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/default";
                case PowerPointSmartArtType.BasicMatrix:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/matrix3";
                case PowerPointSmartArtType.BasicPyramid:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1";
                case PowerPointSmartArtType.BasicRelationship:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/radial1";
                default:
                    return "urn:microsoft.com/office/officeart/2005/8/layout/process1";
            }
        }

        private static string GetSmartArtCategory(PowerPointSmartArtType type) {
            switch (type) {
                case PowerPointSmartArtType.BasicHierarchy: return "hierarchy";
                case PowerPointSmartArtType.BasicCycle: return "cycle";
                case PowerPointSmartArtType.BasicList: return "list";
                case PowerPointSmartArtType.BasicMatrix: return "matrix";
                case PowerPointSmartArtType.BasicPyramid: return "pyramid";
                case PowerPointSmartArtType.BasicRelationship: return "relationship";
                default: return "process";
            }
        }

        private static void PopulateSmartArtColors(DiagramColorsPart part) {
            using MemoryStream stream = new(Encoding.UTF8.GetBytes(
                SmartArtColorsDefinitionXml));
            part.FeedData(stream);
        }

        private static void PopulateSmartArtStyle(DiagramStylePart part) {
            using MemoryStream stream = new(Encoding.UTF8.GetBytes(
                SmartArtStyleDefinitionXml));
            part.FeedData(stream);
        }
    }
}
