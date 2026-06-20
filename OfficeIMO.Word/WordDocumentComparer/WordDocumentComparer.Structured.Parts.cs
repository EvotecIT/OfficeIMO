using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static Dictionary<HeaderPart, string> CreateHeaderPartKeys(MainDocumentPart mainPart) {
            return CreateOrderedHeaderPartKeys(mainPart).ToDictionary(item => item.Key, item => item.Value);
        }

        private static List<KeyValuePair<HeaderPart, string>> CreateOrderedHeaderPartKeys(MainDocumentPart mainPart) {
            var orderedKeys = new List<KeyValuePair<HeaderPart, string>>();
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
                string key = HeaderPartKeyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                keys[headerPart] = key;
                orderedKeys.Add(new KeyValuePair<HeaderPart, string>(headerPart, key));
            }

            return orderedKeys;
        }

        private static Dictionary<FooterPart, string> CreateFooterPartKeys(MainDocumentPart mainPart) {
            return CreateOrderedFooterPartKeys(mainPart).ToDictionary(item => item.Key, item => item.Value);
        }

        private static List<KeyValuePair<FooterPart, string>> CreateOrderedFooterPartKeys(MainDocumentPart mainPart) {
            var orderedKeys = new List<KeyValuePair<FooterPart, string>>();
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
                string key = FooterPartKeyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                keys[footerPart] = key;
                orderedKeys.Add(new KeyValuePair<FooterPart, string>(footerPart, key));
            }

            return orderedKeys;
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

            if (IsTextOnlyHeaderFooterContent(root)) {
                return GetTextOnlyHeaderFooterPartSignature(part, root);
            }

            OpenXmlElement clone = root.CloneNode(true);
            NormalizeHeaderFooterSignatureElement(part, clone);
            return clone.OuterXml;
        }

        private static bool IsTextOnlyHeaderFooterContent(OpenXmlElement root) {
            return !root.Descendants<Table>().Any() &&
                   !root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() &&
                   !root.Descendants<V.ImageData>().Any() &&
                   !root.Descendants<Hyperlink>().Any() &&
                   !root.Descendants<SimpleField>().Any() &&
                   !root.Descendants<FieldCode>().Any();
        }

        private static string GetTextOnlyHeaderFooterPartSignature(OpenXmlPart part, OpenXmlElement root) {
            string[] paragraphs = root.Descendants<Paragraph>()
                .Where(paragraph => paragraph.Ancestors<TableCell>().FirstOrDefault() == null)
                .Select(paragraph => "p:" + GetParagraphMatchText(paragraph, part))
                .ToArray();

            return paragraphs.Length == 0
                ? root.InnerText ?? string.Empty
                : string.Join("\n", paragraphs);
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

        private static List<Footnote> GetReferencedFootnotes(MainDocumentPart mainPart) {
            Dictionary<long, Footnote> footnotesById = mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>()
                .Where(IsVisibleNote)
                .Where(footnote => footnote.Id?.Value != null)
                .ToDictionary(footnote => footnote.Id!.Value, footnote => footnote) ?? new Dictionary<long, Footnote>();
            return GetReferencedNoteIds<FootnoteReference>(mainPart)
                .Where(footnotesById.ContainsKey)
                .Select(noteId => footnotesById[noteId])
                .ToList();
        }

        private static List<Endnote> GetReferencedEndnotes(MainDocumentPart mainPart) {
            Dictionary<long, Endnote> endnotesById = mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>()
                .Where(IsVisibleNote)
                .Where(endnote => endnote.Id?.Value != null)
                .ToDictionary(endnote => endnote.Id!.Value, endnote => endnote) ?? new Dictionary<long, Endnote>();
            return GetReferencedNoteIds<EndnoteReference>(mainPart)
                .Where(endnotesById.ContainsKey)
                .Select(noteId => endnotesById[noteId])
                .ToList();
        }

        private static List<long> GetReferencedNoteIds<TReference>(MainDocumentPart mainPart)
            where TReference : OpenXmlElement {
            var ids = new List<long>();
            var seenIds = new HashSet<long>();
            foreach (OpenXmlElement root in GetNoteReferenceRoots(mainPart)) {
                foreach (TReference reference in EnumerateComparableDescendants(root).OfType<TReference>()) {
                    long? noteId = reference switch {
                        FootnoteReference footnoteReference => footnoteReference.Id?.Value,
                        EndnoteReference endnoteReference => endnoteReference.Id?.Value,
                        _ => null
                    };
                    if (noteId != null && seenIds.Add(noteId.Value)) {
                        ids.Add(noteId.Value);
                    }
                }
            }

            return ids;
        }

        private static IEnumerable<OpenXmlElement> GetNoteReferenceRoots(MainDocumentPart mainPart) {
            if (mainPart.Document != null) {
                yield return mainPart.Document;
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                if (headerPart.Header != null) {
                    yield return headerPart.Header;
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts) {
                if (footerPart.Footer != null) {
                    yield return footerPart.Footer;
                }
            }
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
