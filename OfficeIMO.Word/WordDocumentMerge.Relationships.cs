using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    internal sealed partial class WordDocumentMerge {
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// Points relationship attributes (r:id, r:embed, ...) of imported content at copies of the source parts,
        /// hyperlinks and external targets in the destination container.
        /// </summary>
        private void RemapRelationshipIds(OpenXmlElement element, OpenXmlPartContainer source, OpenXmlPartContainer destination) {
            if (!_relationshipMaps.TryGetValue((source, destination), out Dictionary<string, string>? map)) {
                map = new Dictionary<string, string>(StringComparer.Ordinal);
                _relationshipMaps[(source, destination)] = map;
            }

            RemapRelationshipIds(element, source, destination, map);
        }

        private void RemapRelationshipIds(OpenXmlElement element, OpenXmlPartContainer source, OpenXmlPartContainer destination, Dictionary<string, string> map) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.Equals(attribute.NamespaceUri, RelationshipNamespace, StringComparison.Ordinal) || string.IsNullOrEmpty(attribute.Value)) {
                    continue;
                }

                string newId = EnsureRelationship(attribute.Value!, source, destination, map);
                if (!string.Equals(newId, attribute.Value, StringComparison.Ordinal)) {
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newId));
                }
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                RemapRelationshipIds(child, source, destination, map);
            }
        }

        private string EnsureRelationship(string relationshipId, OpenXmlPartContainer source, OpenXmlPartContainer destination, Dictionary<string, string> map) {
            if (map.TryGetValue(relationshipId, out string? mapped)) {
                return mapped;
            }

            string newId;
            if (source.TryGetPartById(relationshipId, out OpenXmlPart? sourcePart)) {
                OpenXmlPart destinationPart = ClonePart(sourcePart, destination);
                newId = GetOrCreateRelationshipId(destination, destinationPart);
            } else if (source.HyperlinkRelationships.FirstOrDefault(_ => _.Id == relationshipId) is HyperlinkRelationship hyperlink) {
                newId = destination.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
            } else if (source.ExternalRelationships.FirstOrDefault(_ => _.Id == relationshipId) is ExternalRelationship external) {
                newId = destination.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
            } else {
                newId = relationshipId;
            }

            map[relationshipId] = newId;
            return newId;
        }

        private OpenXmlPart ClonePart(OpenXmlPart sourcePart, OpenXmlPartContainer destination) {
            if (_partMap.TryGetValue(sourcePart, out OpenXmlPart? existing)) {
                return existing;
            }

            OpenXmlPart cloned = destination.AddPart(sourcePart);
            _partMap[sourcePart] = cloned;
            CopyPartRelationships(sourcePart, cloned);
            return cloned;
        }

        /// <summary>
        /// Makes sure the copied part has the related parts and links of its source under the same relationship ids,
        /// since its content refers to them by id.
        /// </summary>
        private void CopyPartRelationships(OpenXmlPart sourcePart, OpenXmlPart cloned) {
            foreach (ExternalRelationship external in sourcePart.ExternalRelationships) {
                if (!cloned.ExternalRelationships.Any(_ => _.Id == external.Id)) {
                    cloned.AddExternalRelationship(external.RelationshipType, external.Uri, external.Id);
                }
            }

            foreach (HyperlinkRelationship hyperlink in sourcePart.HyperlinkRelationships) {
                if (!cloned.HyperlinkRelationships.Any(_ => _.Id == hyperlink.Id)) {
                    cloned.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal, hyperlink.Id);
                }
            }

            foreach (IdPartPair child in sourcePart.Parts) {
                if (cloned.TryGetPartById(child.RelationshipId, out OpenXmlPart? copiedChild)) {
                    if (!_partMap.ContainsKey(child.OpenXmlPart)) {
                        _partMap[child.OpenXmlPart] = copiedChild;
                        CopyPartRelationships(child.OpenXmlPart, copiedChild);
                    }

                    continue;
                }

                OpenXmlPart added = cloned.AddPart(child.OpenXmlPart, child.RelationshipId);
                if (!_partMap.ContainsKey(child.OpenXmlPart)) {
                    _partMap[child.OpenXmlPart] = added;
                }

                CopyPartRelationships(child.OpenXmlPart, added);
            }
        }

        private static string GetOrCreateRelationshipId(OpenXmlPartContainer container, OpenXmlPart part) {
            foreach (IdPartPair pair in container.Parts) {
                if (ReferenceEquals(pair.OpenXmlPart, part)) {
                    return pair.RelationshipId;
                }
            }

            return container.CreateRelationshipToPart(part);
        }
    }
}
