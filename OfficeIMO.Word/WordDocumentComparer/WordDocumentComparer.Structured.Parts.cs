using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static Dictionary<WordPartView, string> CreateHeaderPartKeys(WordPartView mainPart) {
            return CreateOrderedHeaderPartKeys(mainPart).ToDictionary(item => item.Key, item => item.Value);
        }

        private static List<KeyValuePair<WordPartView, string>> CreateOrderedHeaderPartKeys(WordPartView mainPart) =>
            CreateOrderedHeaderFooterPartKeys(mainPart, "headerReference", WordPackageView.HeaderRelationshipType, "hdr", HeaderPartKeyPrefix);

        private static Dictionary<WordPartView, string> CreateFooterPartKeys(WordPartView mainPart) {
            return CreateOrderedFooterPartKeys(mainPart).ToDictionary(item => item.Key, item => item.Value);
        }

        private static List<KeyValuePair<WordPartView, string>> CreateOrderedFooterPartKeys(WordPartView mainPart) =>
            CreateOrderedHeaderFooterPartKeys(mainPart, "footerReference", WordPackageView.FooterRelationshipType, "ftr", FooterPartKeyPrefix);

        private static List<KeyValuePair<WordPartView, string>> CreateOrderedHeaderFooterPartKeys(
            WordPartView mainPart,
            string referenceName,
            string relationshipType,
            string rootName,
            string keyPrefix) {
            var orderedKeys = new List<KeyValuePair<WordPartView, string>>();
            var keys = new Dictionary<WordPartView, string>();
            var typeOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
            var seenEffectiveParts = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode reference in DescendantsW(DocumentRoot(mainPart), referenceName)) {
                if (reference.Attribute(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.Relationships, "id") is not string relationshipId) {
                    continue;
                }

                string? type = GetHeaderFooterReferenceType(reference);
                if (!IsHeaderFooterReferenceVisible(mainPart, type, reference)) {
                    continue;
                }

                // The SDK path threw for a reference to a relationship that targets no part.
                WordPartView part = mainPart.GetPartById(relationshipId) ??
                    throw new ArgumentOutOfRangeException("id");
                if (!string.Equals(mainPart.PartRelationshipType(relationshipId), relationshipType, StringComparison.Ordinal) ||
                    keys.ContainsKey(part)) {
                    continue;
                }

                string typeKey = GetHeaderFooterReferenceTypeKey(type);
                if (!seenEffectiveParts.Add(typeKey + ":" + GetHeaderFooterPartSignature(part, part.RootNamed(W, rootName)))) {
                    continue;
                }

                int ordinal = GetAndIncrementOrdinal(typeOrdinals, typeKey);
                string key = keyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                keys[part] = key;
                orderedKeys.Add(new KeyValuePair<WordPartView, string>(part, key));
            }

            return orderedKeys;
        }

        private static string GetHeaderPartKey(IReadOnlyDictionary<WordPartView, string> keys, WordPartView headerPart, int fallbackIndex) {
            return keys.TryGetValue(headerPart, out string? key)
                ? key
                : HeaderPartKeyPrefix + "unreferenced:" + fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetFooterPartKey(IReadOnlyDictionary<WordPartView, string> keys, WordPartView footerPart, int fallbackIndex) {
            return keys.TryGetValue(footerPart, out string? key)
                ? key
                : FooterPartKeyPrefix + "unreferenced:" + fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>The <c>w:type</c> of a header or footer reference: null when absent, a format error when invalid.</summary>
        private static string? GetHeaderFooterReferenceType(OfficeXmlNode reference) =>
            EnumToken(reference.Attribute(W, "type"), "default", "even", "first");

        private static string GetHeaderFooterReferenceTypeKey(string? type) {
            if (type == "first") {
                return "first";
            }

            if (type == "even") {
                return "even";
            }

            return "default";
        }

        private static string GetHeaderFooterPartSignature(WordPartView part, OfficeXmlNode root) {
            if (root.IsNone) {
                return string.Empty;
            }

            if (IsTextOnlyHeaderFooterContent(root)) {
                return GetTextOnlyHeaderFooterPartSignature(part, root);
            }

            return GetElementSignature(root, new HeaderFooterSignatureRules(part));
        }

        private static bool IsTextOnlyHeaderFooterContent(OfficeXmlNode root) {
            foreach (OfficeXmlNode element in TypedDescendants(root)) {
                if (element.Is(W, "tbl") || element.Is(W, "drawing") || element.Is(Vml, "imagedata") ||
                    element.Is(W, "hyperlink") || element.Is(W, "fldSimple") || element.Is(W, "instrText")) {
                    return false;
                }
            }

            return true;
        }

        private static string GetTextOnlyHeaderFooterPartSignature(WordPartView part, OfficeXmlNode root) {
            string[] paragraphs = DescendantsW(root, "p")
                .Where(paragraph => !HasAncestorW(paragraph, "tc"))
                .Select(paragraph => "p:" + GetParagraphMatchText(paragraph, part))
                .ToArray();

            return paragraphs.Length == 0
                ? InnerText(root)
                : string.Join("\n", paragraphs);
        }

        /// <summary>Header and footer markup without revision save ids, relationship ids replaced by what they target.</summary>
        private sealed class HeaderFooterSignatureRules : SignatureRules {
            private readonly WordPartView _part;

            internal HeaderFooterSignatureRules(WordPartView part) => _part = part;

            internal override void Rewrite(OfficeXmlNode element, List<WordXmlAttribute> attributes) {
                for (int index = attributes.Count - 1; index >= 0; index--) {
                    WordXmlAttribute attribute = attributes[index];
                    if (attribute.NamespaceUri == WordprocessingNamespaceUri && attribute.LocalName.StartsWith("rsid", StringComparison.Ordinal)) {
                        attributes.RemoveAt(index);
                    } else if (attribute.NamespaceUri == RelationshipsNamespaceUri &&
                               (attribute.LocalName == "id" || attribute.LocalName == "embed" || attribute.LocalName == "link")) {
                        attributes[index] = new WordXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, GetPartRelationshipSignature(_part, attribute.Value));
                    }
                }

                ResetDrawingProperties(element, attributes);
            }
        }

        private static string GetPartRelationshipSignature(WordPartView part, string relationshipId) {
            if (part.TryGetExternalRelationship(relationshipId, out string relationshipType, out string uri)) {
                return "external:" + relationshipType + ":" + uri;
            }

            string? hyperlinkTarget = part.HyperlinkTarget(relationshipId, out bool isExternal);
            if (hyperlinkTarget != null) {
                return "hyperlink:" + hyperlinkTarget + ":" + isExternal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            WordPartView? relatedPart = part.GetPartById(relationshipId);
            if (relatedPart == null) {
                return "missing:" + relationshipId;
            }

            if (IsImagePartRelationship(part, relationshipId)) {
                ImageFingerprint fingerprint = CreateImageFingerprint(relatedPart);
                return "image:" + relatedPart.ContentType + ":" + fingerprint.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + fingerprint.Sha256;
            }

            return "part:" + relatedPart.ContentType;
        }

        private static bool IsImagePartRelationship(WordPartView part, string relationshipId) =>
            string.Equals(part.PartRelationshipType(relationshipId), WordPackageView.ImageRelationshipType, StringComparison.Ordinal);

        private static bool IsHeaderFooterReferenceVisible(WordPartView mainPart, string? type, OfficeXmlNode reference) {
            if (type == "first") {
                OfficeXmlNode sectionProperties = AncestorW(reference, "sectPr");
                OfficeXmlNode titlePage = sectionProperties.IsNone ? default : sectionProperties.Child(W, "titlePg");
                return !titlePage.IsNone && IsOnOffEnabled(titlePage);
            }

            if (type == "even") {
                OfficeXmlNode settings = SettingsRoot(mainPart);
                return !settings.IsNone && settings.Children(W, "evenAndOddHeaders").ToList().Any(IsOnOffEnabled);
            }

            return true;
        }

        private static int GetAndIncrementOrdinal(Dictionary<string, int> ordinals, string typeKey) {
            ordinals.TryGetValue(typeKey, out int ordinal);
            ordinals[typeKey] = ordinal + 1;
            return ordinal;
        }

        private static bool IsVisibleNote(OfficeXmlNode note) {
            return WordNoteResolver.IsOrdinary(note);
        }

        /// <summary>The <c>w:id</c> of a footnote or endnote as its typed integer reads it, or null.</summary>
        private static long? GetNoteId(OfficeXmlNode note) => ParseInt64(note.Attribute(W, "id"));

        private static string GetNotePartKeyId(OfficeXmlNode note, int fallbackIndex) {
            return GetNoteId(note)?.ToString(System.Globalization.CultureInfo.InvariantCulture) ??
                fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static List<OfficeXmlNode> GetReferencedFootnotes(WordPartView mainPart) =>
            GetReferencedNotes(mainPart, FootnotesRoot(FootnotesPart(mainPart)), "footnote", "footnoteReference");

        private static List<OfficeXmlNode> GetReferencedEndnotes(WordPartView mainPart) =>
            GetReferencedNotes(mainPart, EndnotesRoot(EndnotesPart(mainPart)), "endnote", "endnoteReference");

        private static List<OfficeXmlNode> GetReferencedNotes(WordPartView mainPart, OfficeXmlNode notesRoot, string noteName, string referenceName) {
            var notesById = new Dictionary<long, OfficeXmlNode>();
            if (!notesRoot.IsNone) {
                foreach (OfficeXmlNode note in notesRoot.Children(W, noteName).ToList().Where(IsVisibleNote)) {
                    long? id = GetNoteId(note);
                    if (id == null) {
                        continue;
                    }

                    // The SDK path built this map with ToDictionary, which rejects a repeated identifier.
                    notesById.Add(id.Value, note);
                }
            }

            return GetReferencedNoteIds(mainPart, referenceName)
                .Where(notesById.ContainsKey)
                .Select(noteId => notesById[noteId])
                .ToList();
        }

        private static List<long> GetReferencedNoteIds(WordPartView mainPart, string referenceName) {
            var ids = new List<long>();
            var seenIds = new HashSet<long>();
            foreach (OfficeXmlNode root in GetNoteReferenceRoots(mainPart)) {
                foreach (OfficeXmlNode reference in EnumerateComparableDescendants(root)) {
                    if (!reference.Is(W, referenceName)) {
                        continue;
                    }

                    long? noteId = GetNoteId(reference);
                    if (noteId != null && seenIds.Add(noteId.Value)) {
                        ids.Add(noteId.Value);
                    }
                }
            }

            return ids;
        }

        private static IEnumerable<OfficeXmlNode> GetNoteReferenceRoots(WordPartView mainPart) {
            OfficeXmlNode document = DocumentRoot(mainPart);
            if (!document.IsNone) {
                yield return document;
            }

            foreach (WordPartView headerPart in HeaderParts(mainPart)) {
                OfficeXmlNode header = HeaderRoot(headerPart);
                if (!header.IsNone) {
                    yield return header;
                }
            }

            foreach (WordPartView footerPart in FooterParts(mainPart)) {
                OfficeXmlNode footer = FooterRoot(footerPart);
                if (!footer.IsNone) {
                    yield return footer;
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
