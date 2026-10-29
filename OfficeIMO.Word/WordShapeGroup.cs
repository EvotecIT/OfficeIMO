using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

#nullable enable annotations

namespace OfficeIMO.Word {
    /// <summary>
    /// Describes one bounded DrawingML shape inside a Word shape group.
    /// Coordinates are relative to the group's top-left corner and use physical lengths.
    /// </summary>
    public sealed class WordShapeGroupItem {
        /// <summary>Creates a shape-group item.</summary>
        public WordShapeGroupItem(WordShapeType shapeType, OfficeLength left, OfficeLength top, OfficeLength width, OfficeLength height) {
            ShapeType = shapeType;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        /// <summary>Preset DrawingML geometry.</summary>
        public WordShapeType ShapeType { get; }
        /// <summary>Horizontal position relative to the group.</summary>
        public OfficeLength Left { get; }
        /// <summary>Vertical position relative to the group.</summary>
        public OfficeLength Top { get; }
        /// <summary>Shape width.</summary>
        public OfficeLength Width { get; }
        /// <summary>Shape height.</summary>
        public OfficeLength Height { get; }
        /// <summary>Optional RGB fill, with or without a leading '#'.</summary>
        public string? FillColorHex { get; set; }
        /// <summary>Optional RGB outline, with or without a leading '#'.</summary>
        public string? StrokeColorHex { get; set; }
    }

    /// <summary>
    /// Represents a native Wordprocessing DrawingML group of bounded preset shapes.
    /// This API does not claim arbitrary imported-group editing or rendered Word coordinates.
    /// </summary>
    public sealed class WordShapeGroup : WordElement {
        private readonly WordDrawing _drawing;

        internal WordShapeGroup(WordDocument document, WordParagraph paragraph, Run run, WordDrawing drawing) {
            _document = document;
            _wordParagraph = paragraph;
            _run = run;
            _drawing = drawing;
        }

        internal WordDocument _document;
        internal WordParagraph _wordParagraph;
        internal Run _run;

        /// <summary>Number of preset shape children persisted in the group.</summary>
        public int ChildCount => _drawing.Descendants<Wps.WordprocessingShape>().Count();

        /// <summary>
        /// Reads package-level placement and geometry evidence for this group.
        /// This is not a claim about Word's final rendered coordinates.
        /// </summary>
        public bool TryGetLayoutSnapshot(out WordDrawingLayoutInfo? snapshot) =>
            WordDrawingLayoutReader.TryRead(_drawing, out snapshot);

        /// <summary>Removes the group drawing from its containing paragraph.</summary>
        public void Remove() => _drawing.Remove();

        internal static WordShapeGroup Add(WordParagraph paragraph, IEnumerable<WordShapeGroupItem> items, OfficeLength? left, OfficeLength? top) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            if (items == null) throw new ArgumentNullException(nameof(items));

            List<WordShapeGroupItem> materialized = items.ToList();
            if (materialized.Count < 2) {
                throw new ArgumentException("A shape group requires at least two shapes.", nameof(items));
            }

            bool anchored = left.HasValue || top.HasValue;
            if (left.HasValue != top.HasValue) {
                throw new ArgumentException("Anchored groups require both left and top offsets.", nameof(left));
            }
            if (anchored) {
                if (left!.Value < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(left));
                if (top!.Value < OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(top));
                // wp:posOffset is a signed 32-bit coordinate, unlike a:xfrm's 64-bit coordinates.
                _ = checked((int)left.Value.ToEmus());
                _ = checked((int)top.Value.ToEmus());
            }

            long groupCx = 0;
            long groupCy = 0;
            uint firstDrawingId = WordDrawingIdAllocator.Reserve(paragraph._document!, 2U);
            var children = new List<Wps.WordprocessingShape>(materialized.Count);
            foreach (WordShapeGroupItem item in materialized) {
                if (item.Width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(items), "Shape width must be positive.");
                if (item.Height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(items), "Shape height must be positive.");
                if (item.Left < OfficeLength.Zero || item.Top < OfficeLength.Zero)
                    throw new ArgumentOutOfRangeException(nameof(items), "Child offsets cannot be negative.");
                long x = item.Left.ToEmus();
                long y = item.Top.ToEmus();
                long cx = item.Width.ToEmus();
                long cy = item.Height.ToEmus();
                Wps.WordprocessingShape shape = WordShape.BuildWpsShape(cx, cy, item.ShapeType);
                A.Transform2D transform = shape.Descendants<A.Transform2D>().First();
                transform.Offset = new A.Offset { X = x, Y = y };
                ApplyColors(shape, item.FillColorHex, item.StrokeColorHex);
                children.Add(shape);
                groupCx = Math.Max(groupCx, checked(x + cx));
                groupCy = Math.Max(groupCy, checked(y + cy));
            }

            var transformGroup = new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = groupCx, Cy = groupCy },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = groupCx, Cy = groupCy });
            var group = new Wpg.WordprocessingGroup(
                new Wpg.NonVisualDrawingProperties { Id = firstDrawingId, Name = "Shape Group" },
                new Wpg.NonVisualGroupDrawingShapeProperties(new A.GroupShapeLocks()),
                new Wpg.GroupShapeProperties(transformGroup));
            group.Append(children);

            var graphic = new A.Graphic(
                new A.GraphicData(group) {
                    Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
                });
            OpenXmlElement frame = anchored
                ? WordShape.BuildAnchor(
                    groupCx,
                    groupCy,
                    left!.Value.ToEmus(),
                    top!.Value.ToEmus(),
                    graphic,
                    firstDrawingId + 1U)
                : BuildInline(groupCx, groupCy, graphic, firstDrawingId + 1U);

            var drawing = new WordDrawing(frame);
            Run run = paragraph.VerifyRun();
            run.Append(drawing);
            return new WordShapeGroup(paragraph._document!, paragraph, run, drawing);
        }

        private static DW.Inline BuildInline(long cx, long cy, A.Graphic graphic, uint docPropertiesId) {
            var inline = new DW.Inline {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };
            inline.Append(new DW.Extent { Cx = cx, Cy = cy });
            inline.Append(new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
            inline.Append(new DW.DocProperties { Id = docPropertiesId, Name = "Shape Group" });
            inline.Append(new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }));
            inline.Append(graphic);
            return inline;
        }

        private static void ApplyColors(Wps.WordprocessingShape shape, string? fillColorHex, string? strokeColorHex) {
            Wps.ShapeProperties properties = shape.GetFirstChild<Wps.ShapeProperties>()!;
            if (!string.IsNullOrWhiteSpace(fillColorHex)) {
                properties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = NormalizeRgb(fillColorHex!, nameof(fillColorHex)) }));
            }
            if (!string.IsNullOrWhiteSpace(strokeColorHex)) {
                properties.Append(new A.Outline(
                    new A.SolidFill(new A.RgbColorModelHex { Val = NormalizeRgb(strokeColorHex!, nameof(strokeColorHex)) })));
            }
        }

        private static string NormalizeRgb(string value, string paramName) {
            if (!OfficeIMO.OfficeColor.TryParseHex(value, out OfficeIMO.OfficeColor color)) {
                throw new ArgumentException("Color must be a three-, six-, or eight-digit hexadecimal value.", paramName);
            }
            return color.ToRgbHex();
        }
    }
}
