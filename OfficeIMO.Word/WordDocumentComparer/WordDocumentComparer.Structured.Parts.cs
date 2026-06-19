using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static Dictionary<HeaderPart, string> CreateHeaderPartKeys(MainDocumentPart mainPart) {
            var keys = new Dictionary<HeaderPart, string>();
            var typeOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
            var seenEffectiveParts = new HashSet<string>(StringComparer.Ordinal);
            foreach (HeaderReference reference in mainPart.Document?.Descendants<HeaderReference>() ?? Enumerable.Empty<HeaderReference>()) {
                if (reference.Id?.Value is not string relationshipId) {
                    continue;
                }

                if (!IsHeaderFooterReferenceVisible(mainPart, reference.Type?.Value, reference)) {
                    continue;
                }

                if (mainPart.GetPartById(relationshipId) is not HeaderPart headerPart || keys.ContainsKey(headerPart)) {
                    continue;
                }

                string typeKey = GetHeaderFooterReferenceTypeKey(reference.Type?.Value);
                if (!seenEffectiveParts.Add(typeKey + ":" + GetHeaderFooterPartSignature(headerPart, headerPart.Header))) {
                    continue;
                }

                int ordinal = GetAndIncrementOrdinal(typeOrdinals, typeKey);
                keys[headerPart] = HeaderPartKeyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            return keys;
        }

        private static Dictionary<FooterPart, string> CreateFooterPartKeys(MainDocumentPart mainPart) {
            var keys = new Dictionary<FooterPart, string>();
            var typeOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
            var seenEffectiveParts = new HashSet<string>(StringComparer.Ordinal);
            foreach (FooterReference reference in mainPart.Document?.Descendants<FooterReference>() ?? Enumerable.Empty<FooterReference>()) {
                if (reference.Id?.Value is not string relationshipId) {
                    continue;
                }

                if (!IsHeaderFooterReferenceVisible(mainPart, reference.Type?.Value, reference)) {
                    continue;
                }

                if (mainPart.GetPartById(relationshipId) is not FooterPart footerPart || keys.ContainsKey(footerPart)) {
                    continue;
                }

                string typeKey = GetHeaderFooterReferenceTypeKey(reference.Type?.Value);
                if (!seenEffectiveParts.Add(typeKey + ":" + GetHeaderFooterPartSignature(footerPart, footerPart.Footer))) {
                    continue;
                }

                int ordinal = GetAndIncrementOrdinal(typeOrdinals, typeKey);
                keys[footerPart] = FooterPartKeyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            return keys;
        }

        private static string GetHeaderPartKey(IReadOnlyDictionary<HeaderPart, string> keys, HeaderPart headerPart, int fallbackIndex) {
            return keys.TryGetValue(headerPart, out string? key)
                ? key
                : HeaderPartKeyPrefix + "unreferenced:" + fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetFooterPartKey(IReadOnlyDictionary<FooterPart, string> keys, FooterPart footerPart, int fallbackIndex) {
            return keys.TryGetValue(footerPart, out string? key)
                ? key
                : FooterPartKeyPrefix + "unreferenced:" + fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetHeaderFooterReferenceTypeKey(HeaderFooterValues? type) {
            if (type == HeaderFooterValues.First) {
                return "first";
            }

            if (type == HeaderFooterValues.Even) {
                return "even";
            }

            return "default";
        }

        private static string GetHeaderFooterPartSignature(OpenXmlPart part, OpenXmlElement? root) {
            if (root == null) {
                return string.Empty;
            }

            if (!HasHeaderFooterStructuralContent(root)) {
                return root.InnerText ?? string.Empty;
            }

            OpenXmlElement clone = root.CloneNode(true);
            NormalizeHeaderFooterSignatureElement(part, clone);
            return clone.OuterXml;
        }

        private static bool HasHeaderFooterStructuralContent(OpenXmlElement root) {
            return root.Descendants<Table>().Any() ||
                   root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ||
                   root.Descendants<V.ImageData>().Any() ||
                   root.Descendants<Hyperlink>().Any() ||
                   root.Descendants<SimpleField>().Any() ||
                   root.Descendants<FieldCode>().Any();
        }

        private static void NormalizeHeaderFooterSignatureElement(OpenXmlPart part, OpenXmlElement root) {
            foreach (OpenXmlElement element in new[] { root }.Concat(root.Descendants())) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes().ToList()) {
                    if (IsVolatileHeaderFooterSignatureAttribute(attribute)) {
                        element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                    } else if (IsRelationshipHeaderFooterSignatureAttribute(attribute) && attribute.Value is string relationshipId) {
                        element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, GetPartRelationshipSignature(part, relationshipId)));
                    }
                }

                if (element is DW.DocProperties docProperties) {
                    docProperties.Id = 0U;
                    docProperties.Name = string.Empty;
                } else if (element is PIC.NonVisualDrawingProperties pictureProperties) {
                    pictureProperties.Id = 0U;
                    pictureProperties.Name = string.Empty;
                }
            }
        }

        private static bool IsVolatileHeaderFooterSignatureAttribute(OpenXmlAttribute attribute) {
            return attribute.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
                   attribute.LocalName.StartsWith("rsid", StringComparison.Ordinal);
        }

        private static bool IsRelationshipHeaderFooterSignatureAttribute(OpenXmlAttribute attribute) {
            return attribute.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships" &&
                   (attribute.LocalName == "id" || attribute.LocalName == "embed" || attribute.LocalName == "link");
        }

        private static string GetPartRelationshipSignature(OpenXmlPart part, string relationshipId) {
            ExternalRelationship? externalRelationship = part.ExternalRelationships.FirstOrDefault(item => item.Id == relationshipId);
            if (externalRelationship != null) {
                return "external:" + externalRelationship.RelationshipType + ":" + externalRelationship.Uri;
            }

            HyperlinkRelationship? hyperlinkRelationship = part.HyperlinkRelationships.FirstOrDefault(item => item.Id == relationshipId);
            if (hyperlinkRelationship != null) {
                return "hyperlink:" + hyperlinkRelationship.Uri + ":" + hyperlinkRelationship.IsExternal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            OpenXmlPart relatedPart;
            try {
                relatedPart = part.GetPartById(relationshipId);
            } catch (ArgumentOutOfRangeException) {
                return "missing:" + relationshipId;
            }

            if (relatedPart is ImagePart imagePart) {
                using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
                ImageFingerprint fingerprint = CreateImageFingerprint(stream);
                return "image:" + imagePart.ContentType + ":" + fingerprint.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + fingerprint.Sha256;
            }

            return "part:" + relatedPart.ContentType;
        }

        private static bool IsHeaderFooterReferenceVisible(MainDocumentPart mainPart, HeaderFooterValues? type, OpenXmlElement reference) {
            if (type == HeaderFooterValues.First) {
                SectionProperties? sectionProperties = reference.Ancestors<SectionProperties>().FirstOrDefault();
                TitlePage? titlePage = sectionProperties?.Elements<TitlePage>().FirstOrDefault();
                return titlePage != null && IsOnOffEnabled(titlePage);
            }

            if (type == HeaderFooterValues.Even) {
                Settings? settings = mainPart.DocumentSettingsPart?.Settings;
                return settings?.Elements<EvenAndOddHeaders>().Any(IsOnOffEnabled) == true;
            }

            return true;
        }

        private static bool IsOnOffEnabled(OnOffType onOff) {
            return onOff.Val == null || onOff.Val.Value;
        }

        private static int GetAndIncrementOrdinal(Dictionary<string, int> ordinals, string typeKey) {
            ordinals.TryGetValue(typeKey, out int ordinal);
            ordinals[typeKey] = ordinal + 1;
            return ordinal;
        }

        private static bool IsVisibleNote(Footnote footnote) {
            return footnote.Type == null || footnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static bool IsVisibleNote(Endnote endnote) {
            return endnote.Type == null || endnote.Type.Value == FootnoteEndnoteValues.Normal;
        }

        private static double GetImageSimilarity(ImageSnapshot source, ImageSnapshot target) {
            if (!string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) {
                return 0;
            }

            bool sameVisualSignature = string.Equals(source.VisualSignature, target.VisualSignature, StringComparison.Ordinal);
            if (source.ExternalUri != null || target.ExternalUri != null) {
                if (string.Equals(source.ExternalUri, target.ExternalUri, StringComparison.Ordinal)) {
                    return sameVisualSignature ? 1 : 0.8;
                }

                return sameVisualSignature ? 0.6 : 0;
            }

            if (source.EmbeddedFingerprint != null &&
                target.EmbeddedFingerprint != null &&
                source.EmbeddedFingerprint.Equals(target.EmbeddedFingerprint)) {
                return sameVisualSignature ? 1 : 0.8;
            }

            return sameVisualSignature ? 0.6 : 0;
        }
    }
}
