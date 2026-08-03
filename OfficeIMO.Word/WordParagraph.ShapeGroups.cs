using System.Collections.Generic;
using System.Linq;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;

#nullable enable annotations

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>Gets the native DrawingML shape group hosted by this paragraph, when present.</summary>
        public WordShapeGroup? ShapeGroup {
            get {
                if (_run == null) return null;
                WordDrawing? drawing = _run.ChildElements
                    .OfType<WordDrawing>()
                    .FirstOrDefault(candidate => candidate.Descendants<Wpg.WordprocessingGroup>().Any());
                return drawing == null ? null : new WordShapeGroup(_document, this, _run, drawing);
            }
        }

        /// <summary>Gets whether this paragraph hosts a native DrawingML shape group.</summary>
        public bool IsShapeGroup => ShapeGroup != null;

        /// <summary>Adds an inline native DrawingML group of bounded preset shapes.</summary>
        public WordShapeGroup AddShapeGroup(IEnumerable<WordShapeGroupItem> items) =>
            WordShapeGroup.Add(this, items, null, null);

        /// <summary>Adds a native DrawingML shape group anchored at page-relative point offsets.</summary>
        public WordShapeGroup AddShapeGroup(IEnumerable<WordShapeGroupItem> items, double leftPt, double topPt) =>
            WordShapeGroup.Add(this, items, leftPt, topPt);
    }
}
