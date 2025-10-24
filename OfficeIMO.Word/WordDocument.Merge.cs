using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

            Dictionary<string, string> relationshipIdMap = new();
            Dictionary<OpenXmlPart, OpenXmlPart> partMap = new();

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
                    int? abstractIdValue = abs.AbstractNumberId?.Value;
                    if (abstractIdValue == null) {
                        continue;
                    }

                    int oldAbs = abstractIdValue.Value;
                    int newAbs = GetNextAbstractNumId(destNumbering);
                    abstractMap[oldAbs] = newAbs;
                    var cloneAbs = (AbstractNum)abs.CloneNode(true);
                    cloneAbs.AbstractNumberId = newAbs;
                    destNumbering.Append(cloneAbs);
                }

                foreach (var inst in srcNumbering.Elements<NumberingInstance>()) {
                    int? numberIdValue = inst.NumberID?.Value;
                    if (numberIdValue == null) {
                        continue;
                    }

                    int oldNum = numberIdValue.Value;
                    int newNum = GetNextNumberingId(destNumbering);
                    var cloneInst = (NumberingInstance)inst.CloneNode(true);
                    cloneInst.NumberID = newNum;
                    var absId = cloneInst.GetFirstChild<AbstractNumId>();
                    if (absId != null && absId.Val != null) {
                        int? absIdValue = absId.Val.Value;
                        if (absIdValue != null && abstractMap.TryGetValue(absIdValue.Value, out var mapped)) {
                            absId.Val = mapped;
                        }
                    }
                    destNumbering.Append(cloneInst);
                    numMap[oldNum] = newNum;
                }
            }

            if (srcMain.Document?.Body == null || destMain.Document?.Body == null) return;
            foreach (var element in srcMain.Document.Body.ChildElements) {
                var clone = element.CloneNode(true);
                foreach (var numId in clone.Descendants<NumberingId>()) {
                    int? numberIdValue = numId.Val?.Value;
                    if (numberIdValue != null && numMap.TryGetValue(numberIdValue.Value, out var mapped)) {
                        numId.Val = mapped;
                    }
                }
                RemapRelationshipIds(clone, srcMain, destMain, relationshipIdMap, partMap);
                destMain.Document.Body.Append(clone);
            }
        }

        private static int GetNextAbstractNumId(Numbering numbering) {
            var ids = numbering.ChildElements.OfType<AbstractNum>()
                .Select(n => (int)(n.AbstractNumberId?.Value ?? 0))
                .ToList();
            return ids.Count > 0 ? ids.Max() + 1 : 0;
        }

        private static int GetNextNumberingId(Numbering numbering) {
            var ids = numbering.ChildElements.OfType<NumberingInstance>()
                .Select(n => (int)(n.NumberID?.Value ?? 0))
                .ToList();
            return ids.Count > 0 ? ids.Max() + 1 : 1;
        }

        private static void RemapRelationshipIds(OpenXmlElement element, OpenXmlPartContainer sourceContainer, OpenXmlPartContainer destinationContainer,
            Dictionary<string, string> relationshipIdMap, Dictionary<OpenXmlPart, OpenXmlPart> partMap) {
            const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            foreach (var attribute in element.GetAttributes()) {
                if (!string.Equals(attribute.NamespaceUri, relationshipNamespace, StringComparison.Ordinal) || string.IsNullOrEmpty(attribute.Value)) {
                    continue;
                }

                var newId = EnsureRelationship(attribute.Value!, sourceContainer, destinationContainer, relationshipIdMap, partMap);
                if (!string.Equals(newId, attribute.Value, StringComparison.Ordinal)) {
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newId));
                }
            }

            foreach (var child in element.ChildElements) {
                RemapRelationshipIds(child, sourceContainer, destinationContainer, relationshipIdMap, partMap);
            }
        }

        private static string EnsureRelationship(string relationshipId, OpenXmlPartContainer sourceContainer, OpenXmlPartContainer destinationContainer,
            Dictionary<string, string> relationshipIdMap, Dictionary<OpenXmlPart, OpenXmlPart> partMap) {
            if (relationshipIdMap.TryGetValue(relationshipId, out var mapped)) {
                return mapped;
            }

            OpenXmlPart? sourcePart = TryGetPartById(sourceContainer, relationshipId);
            if (sourcePart != null) {
                var destinationPart = ClonePartRecursive(sourcePart, destinationContainer, partMap);
                var newId = destinationContainer.GetIdOfPart(destinationPart);
                if (string.IsNullOrEmpty(newId)) {
                    newId = destinationContainer.GetIdOfPart(destinationPart);
                }
                relationshipIdMap[relationshipId] = newId!;
                return newId!;
            }

            var hyperlink = sourceContainer.HyperlinkRelationships.FirstOrDefault(h => h.Id == relationshipId);
            if (hyperlink != null) {
                var newId = $"rId{Guid.NewGuid():N}";
                destinationContainer.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal, newId);
                relationshipIdMap[relationshipId] = newId;
                return newId;
            }

            var external = sourceContainer.ExternalRelationships.FirstOrDefault(e => e.Id == relationshipId);
            if (external != null) {
                var newId = $"rId{Guid.NewGuid():N}";
                destinationContainer.AddExternalRelationship(external.RelationshipType, external.Uri, newId);
                relationshipIdMap[relationshipId] = newId;
                return newId;
            }

            relationshipIdMap[relationshipId] = relationshipId;
            return relationshipId;
        }

        private static OpenXmlPart ClonePartRecursive(OpenXmlPart sourcePart, OpenXmlPartContainer destinationContainer, Dictionary<OpenXmlPart, OpenXmlPart> partMap) {
            if (partMap.TryGetValue(sourcePart, out var existingPart)) {
                return existingPart;
            }

            var clonedPart = destinationContainer.AddPart(sourcePart);
            partMap[sourcePart] = clonedPart;

            foreach (var external in sourcePart.ExternalRelationships) {
                clonedPart.AddExternalRelationship(external.RelationshipType, external.Uri);
            }

            foreach (var hyperlink in sourcePart.HyperlinkRelationships) {
                clonedPart.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal);
            }

            foreach (var child in sourcePart.Parts) {
                ClonePartRecursive(child.OpenXmlPart, clonedPart, partMap);
            }

            return clonedPart;
        }

        private static OpenXmlPart? TryGetPartById(OpenXmlPartContainer container, string relationshipId) {
            try {
                return container.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return null;
            }
        }
    }
}
