using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>
    /// Supports merging multiple documents.
    /// </summary>
    public partial class WordDocument {
        /// <summary>
        /// Appends the content of another <see cref="WordDocument"/> to this
        /// document.
        /// </summary>
        /// <param name="source">The document to append.</param>
        public void AppendDocument(WordDocument source) {
            if (source == null) throw new ArgumentNullException(nameof(source));

            var srcMain = source._wordprocessingDocument.MainDocumentPart;
            var destMain = this._wordprocessingDocument.MainDocumentPart;
            if (srcMain == null || destMain == null) return;

            Numbering destNumbering;
            if (destMain.NumberingDefinitionsPart == null) {
                destNumbering = new Numbering();
                destMain.AddNewPart<NumberingDefinitionsPart>().Numbering = destNumbering;
            } else if (destMain.NumberingDefinitionsPart.Numbering == null) {
                destNumbering = destMain.NumberingDefinitionsPart.Numbering = new Numbering();
            } else {
                destNumbering = destMain.NumberingDefinitionsPart.Numbering;
            }

            var srcNumbering = srcMain.NumberingDefinitionsPart?.Numbering;
            Dictionary<int, int> numMap = new();
            Dictionary<int, int> abstractMap = new();
            if (srcNumbering != null) {
                foreach (var abs in srcNumbering.Elements<AbstractNum>()) {
                    int oldAbs = (int)abs.AbstractNumberId.Value;
                    int newAbs = GetNextAbstractNumId(destNumbering);
                    abstractMap[oldAbs] = newAbs;
                    var cloneAbs = (AbstractNum)abs.CloneNode(true);
                    cloneAbs.AbstractNumberId = newAbs;
                    destNumbering.Append(cloneAbs);
                }

                foreach (var inst in srcNumbering.Elements<NumberingInstance>()) {
                    int oldNum = (int)inst.NumberID.Value;
                    int newNum = GetNextNumberingId(destNumbering);
                    var cloneInst = (NumberingInstance)inst.CloneNode(true);
                    cloneInst.NumberID = newNum;
                    var absId = cloneInst.GetFirstChild<AbstractNumId>();
                    if (absId != null && abstractMap.TryGetValue((int)absId.Val.Value, out var mapped)) {
                        absId.Val = mapped;
                    }
                    destNumbering.Append(cloneInst);
                    numMap[oldNum] = newNum;
                }
            }

            foreach (var element in srcMain.Document.Body.ChildElements) {
                var clone = element.CloneNode(true);
                foreach (var numId in clone.Descendants<NumberingId>()) {
                    if (numMap.TryGetValue((int)numId.Val.Value, out var mapped)) {
                        numId.Val = mapped;
                    }
                }
                destMain.Document.Body.Append(clone);
            }
        }

        private static int GetNextAbstractNumId(Numbering numbering) {
            var ids = numbering.ChildElements.OfType<AbstractNum>()
                .Select(n => (int)n.AbstractNumberId.Value)
                .ToList();
            return ids.Count > 0 ? ids.Max() + 1 : 0;
        }

        private static int GetNextNumberingId(Numbering numbering) {
            var ids = numbering.ChildElements.OfType<NumberingInstance>()
                .Select(n => (int)n.NumberID.Value)
                .ToList();
            return ids.Count > 0 ? ids.Max() + 1 : 1;
        }
    }
}
