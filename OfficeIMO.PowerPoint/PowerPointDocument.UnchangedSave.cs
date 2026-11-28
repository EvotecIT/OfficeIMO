using System.Globalization;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Saving a presentation whose SDK package was never opened writes the loaded package again without the SDK: every
    /// part is kept as it was, and docProps/app.xml and docProps/core.xml get the values a save through the SDK gives
    /// them. Anything this path does not cover opens the SDK and saves as before.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private const string NotesSlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
        private const string NotesSlideRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesSlide";
        private const string VbaProjectRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        // Kept in their own type, so the document's static constructor, which the SDK-bound statics share, doesn't carry them.
        private static class UnchangedSaveNames {
            internal static readonly HashSet<string> SlideSizeTypes = new(StringComparer.Ordinal) {
                "screen4x3", "letter", "A4", "35mm", "overhead", "banner", "custom", "ledger", "A3", "B4ISO", "B5ISO", "B4JIS", "B5JIS",
                "hagakiCard", "screen16x9", "screen16x10"
            };
        }

        /// <summary>
        /// Writes the package held by the owned engine with refreshed document properties. A subtype change or VBA
        /// removal applies to the saved copy, while the open document retains its source subtype and project.
        /// </summary>
        private bool TryCreateOwnedPackageBytes(bool keepsSubtype, string? destinationContentType, bool preserveVbaProject, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            if (!_sdkDeferred || IsSdkLoaded || !keepsSubtype && destinationContentType == null
                || _packageStream is not { CanRead: true, CanSeek: true } packageStream) {
                return false;
            }

            UnchangedSaveValues values;
            packageStream.Position = 0;
            // The package the load opened reads the bytes the document still holds; edits since are in the pending package.
            OfficeOpcPackage? loaded = _loadedPackage;
            OfficeOpcPackage package = loaded ?? OfficeOpcPackage.Open(packageStream, leaveOpen: true);
            try {
                if (!TryReadUnchangedSaveValues(package, _maximumSlideXmlCharacters, _pendingPackage,
                    _packageNotesAdded, out values)) return false;
            } finally {
                if (loaded == null) package.Dispose();
            }

            // Slide edits made without the SDK are saved with the properties, in one write.
            OfficeOpcEditablePackage? pending = _pendingPackage;
            packageStream.Position = 0;
            OfficeOpcEditablePackage editable = pending ?? OfficeOpcEditablePackage.Open(packageStream, leaveOpen: true);
            bool kept = false;
            try {
                bool stripVba = values.HasVbaProject && !preserveVbaProject;
                bool projectOutput = !keepsSubtype || stripVba;
                if (stripVba && destinationContentType == null) return false;
                bool validateSubtype = projectOutput || !IsLoadedStandardPresentation;
                if (validateSubtype && !CanSaveCurrentSubtype(editable, projectOutput, keepsSubtype ? null : values.MainPart)) return false;
                PackageSaveProjection? projection = null;
                if (projectOutput && !TryPlanPackageSaveProjection(editable, values.MainPart, keepsSubtype ? null : destinationContentType,
                    stripVba, out projection)) return false;
                if (!TryGetSavePropertyRoots(editable, values.CorePart, values.ApplicationPart, validateSubtype,
                    out OfficeXmlNode core, out OfficeXmlNode application)) return false;

                EnsureApplicationProperty(application, "TotalTime", "0");
                EnsureApplicationProperty(application, "Application", "Microsoft Office PowerPoint");
                SetApplicationProperty(application, "PresentationFormat", values.PresentationFormat);
                SetApplicationProperty(application, "Slides", values.Slides.ToString(CultureInfo.InvariantCulture));
                SetApplicationProperty(application, "Notes", values.Notes.ToString(CultureInfo.InvariantCulture));
                SetApplicationProperty(application, "HiddenSlides", values.HiddenSlides.ToString(CultureInfo.InvariantCulture));
                DateTime modified = DateTime.UtcNow;
                SetCoreProperty(core, OfficeXmlNamespaces.DublinCoreTerms, "modified", FormatPackageDate(modified));

                // The package is written into pooled chunks where the runtime has them, so the array the caller owns is
                // the only copy.
#if NET8_0_OR_GREATER
                using (var saved = new OfficeIMO.Core.Internal.OfficePooledWriteStream()) {
#else
                using (var saved = new MemoryStream(checked((int)Math.Min(int.MaxValue / 2, packageStream.Length + 4096)))) {
#endif
                    if (projection == null) editable.Save(saved);
                    else editable.Save(saved, CreatePackageSaveProjection(editable, projection));
                    bytes = saved.ToArray();
                }
                // The document now holds what it saved, as a save through the SDK leaves it: it keeps the package it
                // wrote, with the refreshed properties, and writes it out again only for whatever reads its bytes next.
                if (pending == null) {
                    _pendingPackage = editable;
                    _pendingSource = packageStream;
                }
                kept = true;
                _pendingEdits = true;
                if (_accessMode == DocumentAccessMode.ReadWrite && (_properties != null || _applicationProperties != null)) {
                    RememberAssignedPackageDate("modified", modified);
                }
            } finally {
                if (!kept && pending == null) editable.Dispose();
            }
            return true;
        }

        private static bool TryGetSavePropertyRoots(OfficeOpcEditablePackage package, string? corePart, string? applicationPart,
            bool validateCore, out OfficeXmlNode core, out OfficeXmlNode application) {
            core = default;
            application = default;
            if (applicationPart != null) {
                OfficeXmlTree tree = package.GetXml(applicationPart);
                application = new OfficeXmlNode(tree, tree.Root);
                if (!application.Is(OfficeXmlNamespaces.ExtendedProperties, "Properties")) return false;
            } else if (package.PartExists(ApplicationPropertiesPartName)) {
                // An unrelated part occupies the usual name. Let the SDK choose a new one.
                return false;
            }
            if (corePart != null) {
                OfficeXmlTree tree = package.GetXml(corePart);
                core = new OfficeXmlNode(tree, tree.Root);
                if (!core.Is(OfficeXmlNamespaces.CoreProperties, "coreProperties") ||
                    ((validateCore || core.Child(OfficeXmlNamespaces.DublinCoreTerms, "modified").IsNone) && !IsValidCoreProperties(core))) {
                    return false;
                }
            }

            if (application.IsNone) application = CreateApplicationPropertiesPart(package);
            if (core.IsNone) core = CreateCorePropertiesPart(package);
            return true;
        }

        // System.IO.Packaging writes package dates in UTC with seven fractional digits.
        private static string FormatPackageDate(DateTime value) =>
            value.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", CultureInfo.InvariantCulture);

        private readonly struct UnchangedSaveValues {
            internal UnchangedSaveValues(string mainPart, string? applicationPart, string? corePart, int slides, int notes,
                int hiddenSlides, string presentationFormat, bool hasVbaProject) {
                MainPart = mainPart;
                ApplicationPart = applicationPart;
                CorePart = corePart;
                Slides = slides;
                Notes = notes;
                HiddenSlides = hiddenSlides;
                PresentationFormat = presentationFormat;
                HasVbaProject = hasVbaProject;
            }

            internal string MainPart { get; }
            internal string? ApplicationPart { get; }
            internal string? CorePart { get; }
            internal int Slides { get; }
            internal int Notes { get; }
            internal int HiddenSlides { get; }
            internal string PresentationFormat { get; }
            internal bool HasVbaProject { get; }
        }

        /// <summary>
        /// Reads what a save through the SDK would write into the document properties, following
        /// <see cref="PowerPointUtils.UpdateDocumentProperties"/>. Returns false for anything that save would also change
        /// elsewhere, including footer definitions and legacy slide visibility.
        /// </summary>
        private static bool TryReadUnchangedSaveValues(OfficeOpcPackage package, long maximumSlideXmlCharacters,
            OfficeOpcEditablePackage? parsed, bool notesAdded, out UnchangedSaveValues values) {
            values = default;
            string? main = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
            if (main == null) return false;
            string? application, core;
            bool validProperties = parsed != null
                ? TryReadPropertyPartNames(parsed, out application, out core)
                : TryReadPropertyPartNames(package, out application, out core);
            if (!validProperties) return false;
            // A dangling relationship or a part beyond the load's limit belongs to the compatibility path.
            if (!CanReadSavePropertyPart(package, parsed, application, maximumSlideXmlCharacters)
                || !CanReadSavePropertyPart(package, parsed, core, maximumSlideXmlCharacters)) return false;

            // Slides removed or moved without the SDK changed the slide list in the pending package, which is read instead.
            OfficeXmlTree? presentationTree = parsed?.GetParsedXml(main);
            var slideIds = new List<string>();
            string presentationFormat = "Widescreen";
            if (presentationTree != null) {
                if (!TryReadPresentation(presentationTree, slideIds, ref presentationFormat)) return false;
            } else {
                OfficeXmlPullReader reader = package.OpenXml(main, out OfficePooledBytes content, processMarkupCompatibility: false);
                using (content) {
                    // Slides are the entries of the first slide list under the root, as the SDK reads them; footer
                    // definitions are extensions in the root's extension list.
                    bool inSlideList = false, slideListSeen = false, inExtensionList = false;
                    while (reader.Read()) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        if (reader.Depth == 2) {
                            inSlideList = !slideListSeen && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldIdLst");
                            slideListSeen |= inSlideList;
                            inExtensionList = reader.IsElement(OfficeXmlNamespaces.PresentationML, "extLst");
                        }
                        if (reader.Depth == 3 && inSlideList && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldId")) {
                            // A legacy visibility attribute makes the SDK save rewrite the slide list and the slide.
                            if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "show", out _)) return false;
                            if (!reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) return false;
                            slideIds.Add(id);
                        } else if (reader.Depth == 2 && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldSz")) {
                            if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "type", out string sizeType)) {
                                if (!UnchangedSaveNames.SlideSizeTypes.Contains(sizeType)) return false;
                                presentationFormat = sizeType == "screen4x3" ? "Standard" : sizeType == "custom" ? "Custom" : "Widescreen";
                            }
                        } else if (reader.Depth == 3 && inExtensionList && reader.IsElement(OfficeXmlNamespaces.PresentationML, "ext")
                            && reader.TryGetAttribute(OfficeXmlNamespaces.None, "uri", out string uri)
                            && string.Equals(uri, FooterDefinitionsUri, StringComparison.Ordinal)) {
                            // Footer definitions are written again on every save.
                            return false;
                        }
                    }
                }
            }

            int hidden = 0;
            // The presentation's relationships from the pending package, by id, read once rather than per slide.
            Dictionary<string, string?>? pendingTargets = null;
            if (presentationTree != null) {
                pendingTargets = new Dictionary<string, string?>(StringComparer.Ordinal);
                foreach ((string id, string type, string? target, bool external) in MainRelationshipsById(parsed!, main)) {
                    if (!pendingTargets.ContainsKey(id)) pendingTargets.Add(id, external ? null : target);
                }
            }
            foreach (string id in slideIds) {
                string? slide;
                if (pendingTargets != null) {
                    slide = pendingTargets.TryGetValue(id, out string? target) ? target : null;
                    if (slide == null || !parsed!.PartExists(slide)) return false;
                } else {
                    OfficeOpcRelationship? relationship = package.FindRelationship(main, id);
                    slide = relationship?.TargetPartName;
                    if (slide == null || relationship!.IsExternal || !package.PartExists(slide)) return false;
                }
                // A slide a replacement already parsed is read from its tree; text edits leave its visibility as it was.
                bool? show;
                // A slide added without the SDK, parsed or not, is only in the pending package, even where it took the
                // name of a slide removed before it.
                if ((parsed?.GetParsedXml(slide) ?? (parsed != null && !parsed.IsStoredPart(slide) ? parsed.GetXml(slide) : null)) is { } tree) {
                    // A slide added without the SDK is not in the loaded bytes, and the model wrote it small.
                    if (parsed!.IsStoredPart(slide) && package.GetPartLength(slide) > SlidePartLimit(maximumSlideXmlCharacters)) return false;
                    if (!TryParseShow(tree.GetAttribute(tree.Root, OfficeXmlNamespaces.None, "show"), out show)) return false;
                } else if (!TryReadSlideShow(package, slide, maximumSlideXmlCharacters, out show)) {
                    return false;
                }
                if (show == false) hidden++;
            }

            // Notes count slide parts, each once, as the SDK's SlideParts does. A notes slide carries its content type, or
            // the SDK does not open the package, so without one declared no slide's relationships need reading.
            var slideParts = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            bool notesDeclared = notesAdded || package.DeclaresContentType(NotesSlideContentType);
            int notes = 0;
            bool hasVbaProject = false;
            bool pendingRelationships = parsed?.GetParsedXml(OfficeOpcPartNames.GetRelationshipsPartName(main)) != null;
            foreach ((string type, string? target, bool external) in MainRelationships(package, parsed, pendingRelationships, main)) {
                if (string.Equals(type, VbaProjectRelationshipType, StringComparison.Ordinal)) hasVbaProject = true;
                if (external || !string.Equals(type, OfficeOpcPackage.SlideRelationshipType, StringComparison.Ordinal)
                    || target is not { } slide || !slideParts.Add(OfficeOpcPackage.GetPartIdentity(slide))) {
                    continue;
                }
                // A slide still listed keeps its relationships as loaded; a removed slide's notes left with it.
                if (!notesDeclared) continue;
                // A slide whose relationships changed without the SDK, or that was added that way, is read from the
                // pending package; the others keep their relationships as loaded.
                string slideRelationships = OfficeOpcPartNames.GetRelationshipsPartName(slide);
                string? notesPart = parsed != null && (!package.PartExists(slide) || parsed.GetParsedXml(slideRelationships) != null)
                    ? FirstRelatedPart(parsed, slide, NotesSlideRelationshipType)
                    : package.GetRelatedPartName(slide, NotesSlideRelationshipType);
                if (notesPart != null && (parsed?.PartExists(notesPart) ?? package.PartExists(notesPart))) notes++;
            }

            values = new UnchangedSaveValues(main, application, core, slideIds.Count, notes, hidden, presentationFormat, hasVbaProject);
            return true;
        }

        private static bool CanReadSavePropertyPart(OfficeOpcPackage source, OfficeOpcEditablePackage? pending, string? part, long limit) {
            if (part == null) return true;
            if (!(pending?.PartExists(part) ?? source.PartExists(part))) return false;
            return pending?.IsStoredPart(part) == false || source.GetPartLength(part) <= limit;
        }

        private static string? FirstRelatedPart(OfficeOpcEditablePackage package, string? source, string type) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source);
            if (!package.PartExists(relationshipsPart)) return null;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                if (OfficeXmlNamespaces.NormalizeRelationshipType(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty) != type) continue;
                return OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
            }
            return null;
        }

        // The pending package's presentation, read as the pull reader reads the loaded one.
        private static bool TryReadPresentation(OfficeXmlTree presentation, List<string> slideIds, ref string presentationFormat) {
            int root = presentation.Root;
            if (!presentation.IsElement(root, OfficeXmlNamespaces.PresentationML, "presentation")) return false;
            int list = presentation.FindChild(root, OfficeXmlNamespaces.PresentationML, "sldIdLst");
            if (list != OfficeXmlTree.NoNode) {
                for (int entry = presentation.FindChild(list, OfficeXmlNamespaces.PresentationML, "sldId"); entry != OfficeXmlTree.NoNode;
                     entry = presentation.FindNextSibling(entry, OfficeXmlNamespaces.PresentationML, "sldId")) {
                    if (presentation.TryGetAttribute(entry, OfficeXmlNamespaces.None, "show", out _)) return false;
                    if (!presentation.TryGetAttribute(entry, OfficeXmlNamespaces.Relationships, "id", out string id)) return false;
                    slideIds.Add(id);
                }
            }
            for (int child = presentation.FirstChildOf(root); child != OfficeXmlTree.NoNode; child = presentation.NextSiblingOf(child)) {
                if (presentation.IsElement(child, OfficeXmlNamespaces.PresentationML, "sldSz")) {
                    if (presentation.TryGetAttribute(child, OfficeXmlNamespaces.None, "type", out string sizeType)) {
                        if (!UnchangedSaveNames.SlideSizeTypes.Contains(sizeType)) return false;
                        presentationFormat = sizeType == "screen4x3" ? "Standard" : sizeType == "custom" ? "Custom" : "Widescreen";
                    }
                } else if (presentation.IsElement(child, OfficeXmlNamespaces.PresentationML, "extLst")) {
                    for (int extension = presentation.FindChild(child, OfficeXmlNamespaces.PresentationML, "ext"); extension != OfficeXmlTree.NoNode;
                         extension = presentation.FindNextSibling(extension, OfficeXmlNamespaces.PresentationML, "ext")) {
                        if (string.Equals(presentation.GetAttribute(extension, OfficeXmlNamespaces.None, "uri"), FooterDefinitionsUri, StringComparison.Ordinal)) return false;
                    }
                }
            }
            return true;
        }

        // The pending package's relationships of the presentation part, with their ids.
        private static IEnumerable<(string Id, string Type, string? Target, bool External)> MainRelationshipsById(OfficeOpcEditablePackage parsed, string main) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(main);
            if (!parsed.PartExists(relationshipsPart)) yield break;
            OfficeXmlTree relationships = parsed.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string? id = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                if (id == null) continue;
                bool external = string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                yield return (id, string.Empty, external ? null : OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target")), external);
            }
        }

        // The presentation part's relationships may change independently of the slide list.
        private static IEnumerable<(string Type, string? Target, bool External)> MainRelationships(OfficeOpcPackage package,
            OfficeOpcEditablePackage? parsed, bool pending, string main) {
            if (!pending) {
                foreach (OfficeOpcRelationship relationship in package.GetRelationships(main)) {
                    yield return (relationship.Type, relationship.TargetPartName, relationship.IsExternal);
                }
                yield break;
            }
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(main);
            if (!parsed!.PartExists(relationshipsPart)) yield break;
            OfficeXmlTree relationships = parsed.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string type = OfficeXmlNamespaces.NormalizeRelationshipType(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty);
                bool external = string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                yield return (type, external ? null : OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target")), external);
            }
        }

        // The slide root's show attribute, as PowerPointXmlReader.ReadSlideShow reads it for the SDK save. A part over the
        // size limit, a value XML Schema does not read as a boolean, or markup the pull reader rejects leaves the
        // answer to that save, which fails or succeeds as before.
        private static long SlidePartLimit(long maximumSlideXmlCharacters) => maximumSlideXmlCharacters > 0
            ? Math.Min(PowerPointXmlReader.MaximumPackageXmlCharacters, maximumSlideXmlCharacters)
            : PowerPointXmlReader.MaximumPackageXmlCharacters;

        private static bool TryReadSlideShow(OfficeOpcPackage package, string slide, long maximumSlideXmlCharacters, out bool? show) {
            show = null;
            if (package.GetPartLength(slide) > SlidePartLimit(maximumSlideXmlCharacters)) return false;
            try {
                OfficeXmlPullReader reader = package.OpenXml(slide, out OfficePooledBytes content, processMarkupCompatibility: false);
                using (content) {
                    while (reader.Read()) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                        return TryParseShow(reader.TryGetAttribute(OfficeXmlNamespaces.None, "show", out string value) ? value : null, out show);
                    }
                }
                return false;
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return false;
            }
        }

        // An absent attribute is no answer; a value XML Schema does not read as a boolean leaves it to the SDK save.
        private static bool TryParseShow(string? value, out bool? show) {
            show = null;
            if (value == null) return true;
            switch (value.Trim(' ', '\t', '\r', '\n')) {
                case "true": case "1": show = true; return true;
                case "false": case "0": show = false; return true;
                default: return false;
            }
        }
    }
}
