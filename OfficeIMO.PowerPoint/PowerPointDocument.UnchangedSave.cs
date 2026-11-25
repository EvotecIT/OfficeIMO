using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Saving a presentation whose SDK package was never opened writes the loaded package again without the SDK: every
    /// part is kept as it was, and docProps/app.xml and docProps/core.xml get the values a save through the SDK gives
    /// them. Anything this path does not cover opens the SDK and saves as before.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string NotesSlideRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesSlide";
        private const string VbaProjectRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        private static readonly string[] ApplicationCounters = { "TotalTime", "Application", "PresentationFormat", "Slides", "Notes", "HiddenSlides" };
        private static readonly HashSet<string> SlideSizeTypes = new(StringComparer.Ordinal) {
            "screen4x3", "letter", "A4", "35mm", "overhead", "banner", "custom", "ledger", "A3", "B4ISO", "B5ISO", "B4JIS", "B5JIS",
            "hagakiCard", "screen16x9", "screen16x10"
        };

        /// <summary>
        /// Writes the loaded package with refreshed document properties when the SDK was never opened and the save
        /// needs nothing else; returns false, having changed nothing, when it must go through the SDK.
        /// </summary>
        private bool TryCreateUnchangedPackageBytes(PresentationDocumentType? destinationType, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (!_sdkDeferred || _sdkDocument != null || destinationType is { } type && type != PresentationDocumentType.Presentation
                || _packageStream is not { CanRead: true, CanSeek: true } packageStream) {
                return false;
            }

            UnchangedSaveValues values;
            packageStream.Position = 0;
            using (OfficeOpcPackage package = OfficeOpcPackage.Open(packageStream, leaveOpen: true)) {
                if (!TryReadUnchangedSaveValues(package, _maximumSlideXmlCharacters, out values)) return false;
            }

            packageStream.Position = 0;
            using (OfficeOpcEditablePackage editable = OfficeOpcEditablePackage.Open(packageStream, leaveOpen: true)) {
                OfficeXmlTree application = editable.GetXml(values.ApplicationPart);
                SetChildText(application, "Slides", values.Slides.ToString(CultureInfo.InvariantCulture));
                SetChildText(application, "Notes", values.Notes.ToString(CultureInfo.InvariantCulture));
                SetChildText(application, "HiddenSlides", values.HiddenSlides.ToString(CultureInfo.InvariantCulture));
                SetChildText(application, "PresentationFormat", values.PresentationFormat);

                OfficeXmlTree core = editable.GetXml(values.CorePart);
                int modified = core.FindChild(core.Root, OfficeXmlNamespaces.DublinCoreTerms, "modified");
                core.SetInnerText(modified, FormatPackageDate(DateTime.UtcNow));

                using var output = new MemoryStream();
                editable.Save(output);
                bytes = output.ToArray();
            }
            // The document now holds what it saved, as a save through the SDK leaves it: an SDK opened later reads the
            // refreshed properties.
            _packageStream = OfficeIMO.Core.Internal.OfficePackageLoadSource.CreatePackageStream(bytes, _accessMode);
            packageStream.Dispose();
            return true;
        }

        // A save through the SDK writes the element text even when it is unchanged; here an unchanged value keeps its bytes.
        private static void SetChildText(OfficeXmlTree tree, string localName, string text) {
            int child = tree.FindChild(tree.Root, OfficeXmlNamespaces.ExtendedProperties, localName);
            if (!string.Equals(tree.GetInnerText(child), text, StringComparison.Ordinal)) tree.SetInnerText(child, text);
        }

        // System.IO.Packaging writes package dates in UTC with seven fractional digits.
        private static string FormatPackageDate(DateTime value) =>
            value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);

        private readonly struct UnchangedSaveValues {
            internal UnchangedSaveValues(string applicationPart, string corePart, int slides, int notes, int hiddenSlides, string presentationFormat) {
                ApplicationPart = applicationPart;
                CorePart = corePart;
                Slides = slides;
                Notes = notes;
                HiddenSlides = hiddenSlides;
                PresentationFormat = presentationFormat;
            }

            internal string ApplicationPart { get; }
            internal string CorePart { get; }
            internal int Slides { get; }
            internal int Notes { get; }
            internal int HiddenSlides { get; }
            internal string PresentationFormat { get; }
        }

        /// <summary>
        /// Reads what a save through the SDK would write into the document properties, following
        /// <see cref="PowerPointUtils.UpdateDocumentProperties"/>. Returns false for anything that save would also change
        /// elsewhere or create: footer definitions, legacy slide visibility, a VBA project, or missing property parts.
        /// </summary>
        private static bool TryReadUnchangedSaveValues(OfficeOpcPackage package, long maximumSlideXmlCharacters, out UnchangedSaveValues values) {
            values = default;
            string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
            string? application = package.GetRelatedPartName(null, ExtendedPropertiesRelationshipType);
            string? core = package.GetRelatedPartName(null, CorePropertiesRelationshipType);
            if (main == null || application == null || core == null || !package.PartExists(application) || !package.PartExists(core)) return false;
            // The SDK save reads both property parts under the part size limit; a part over it takes that path.
            if (package.GetPartLength(application) > maximumSlideXmlCharacters || package.GetPartLength(core) > maximumSlideXmlCharacters) return false;
            if (!HasApplicationCounters(package, application) || !HasModifiedDate(package, core)) return false;

            var slideIds = new List<string>();
            string presentationFormat = "Widescreen";
            OfficeXmlPullReader reader = package.OpenXml(main, out OfficePooledBytes content);
            using (content) {
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (reader.Depth == 3 && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldId")) {
                        // A legacy visibility attribute makes the SDK save rewrite the slide list and the slide.
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "show", out _)) return false;
                        if (!reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) return false;
                        slideIds.Add(id);
                    } else if (reader.Depth == 2 && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldSz")) {
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "type", out string sizeType)) {
                            if (!SlideSizeTypes.Contains(sizeType)) return false;
                            presentationFormat = sizeType == "screen4x3" ? "Standard" : sizeType == "custom" ? "Custom" : "Widescreen";
                        }
                    } else if (reader.Depth == 3 && reader.IsElement(OfficeXmlNamespaces.PresentationML, "ext")
                        && reader.TryGetAttribute(OfficeXmlNamespaces.None, "uri", out string uri)
                        && string.Equals(uri, FooterDefinitionsUri, StringComparison.Ordinal)) {
                        // Footer definitions are written again on every save.
                        return false;
                    }
                }
            }

            int hidden = 0;
            foreach (string id in slideIds) {
                OfficeOpcRelationship? relationship = package.FindRelationship(main, id);
                if (relationship?.TargetPartName is not { } slide || relationship.IsExternal || !package.PartExists(slide)) return false;
                if (!TryReadSlideShow(package, slide, maximumSlideXmlCharacters, out bool? show)) return false;
                if (show == false) hidden++;
            }

            // Notes count slide parts, each once, as the SDK's SlideParts does.
            var slideParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int notes = 0;
            foreach (OfficeOpcRelationship relationship in package.GetRelationships(main)) {
                if (string.Equals(relationship.Type, VbaProjectRelationshipType, StringComparison.Ordinal)) return false;
                if (relationship.IsExternal || !string.Equals(relationship.Type, OfficeOpcPackage.SlideRelationshipType, StringComparison.Ordinal)
                    || relationship.TargetPartName is not { } slide || !slideParts.Add(OfficeOpcPackage.GetPartIdentity(slide))) {
                    continue;
                }
                if (package.GetRelatedPartName(slide, NotesSlideRelationshipType) is { } notesPart && package.PartExists(notesPart)) notes++;
            }

            values = new UnchangedSaveValues(application, core, slideIds.Count, notes, hidden, presentationFormat);
            return true;
        }

        // The slide root's show attribute, as PowerPointXmlReader.ReadSlideShow reads it for the SDK save. A part over the
        // size limit, a value XML Schema does not read as a boolean, or markup the pull reader rejects leaves the
        // answer to that save, which fails or succeeds as before.
        private static bool TryReadSlideShow(OfficeOpcPackage package, string slide, long maximumSlideXmlCharacters, out bool? show) {
            show = null;
            long limit = maximumSlideXmlCharacters > 0
                ? Math.Min(PowerPointXmlReader.MaximumPackageXmlCharacters, maximumSlideXmlCharacters)
                : PowerPointXmlReader.MaximumPackageXmlCharacters;
            if (package.GetPartLength(slide) > limit) return false;
            try {
                OfficeXmlPullReader reader = package.OpenXml(slide, out OfficePooledBytes content, processMarkupCompatibility: false);
                using (content) {
                    while (reader.Read()) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        if (!reader.TryGetAttribute(OfficeXmlNamespaces.None, "show", out string value)) return true;
                        switch (value.Trim(' ', '\t', '\r', '\n')) {
                            case "true": case "1": show = true; return true;
                            case "false": case "0": show = false; return true;
                            default: return false;
                        }
                    }
                }
                return false;
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return false;
            }
        }

        // The SDK save adds any of these that are missing; this path only replaces their text.
        private static bool HasApplicationCounters(OfficeOpcPackage package, string application) {
            var found = new HashSet<string>(StringComparer.Ordinal);
            OfficeXmlPullReader reader = package.OpenXml(application, out OfficePooledBytes content);
            using (content) {
                while (reader.Read()) {
                    if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                    if (reader.Depth == 1 && !reader.IsElement(OfficeXmlNamespaces.ExtendedProperties, "Properties")) return false;
                    if (reader.Depth == 2 && reader.NamespaceId == OfficeXmlNamespaces.ExtendedProperties) found.Add(reader.LocalName);
                }
            }
            foreach (string counter in ApplicationCounters) {
                if (!found.Contains(counter)) return false;
            }
            return true;
        }

        private static bool HasModifiedDate(OfficeOpcPackage package, string core) {
            OfficeXmlPullReader reader = package.OpenXml(core, out OfficePooledBytes content);
            using (content) {
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.Depth == 2
                        && reader.IsElement(OfficeXmlNamespaces.DublinCoreTerms, "modified")) {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}
