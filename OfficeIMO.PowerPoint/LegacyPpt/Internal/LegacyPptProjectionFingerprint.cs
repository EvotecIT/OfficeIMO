using System.Collections.ObjectModel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Retains a global package guard plus independent normalized slide and master-theme guards. Existing projected
    /// slides can be reordered or removed, and mapped master themes can be edited, while unsupported mutations to a
    /// retained slide or shared package part are rejected.
    /// </summary>
    internal sealed class LegacyPptProjectionFingerprint {
        private delegate bool TryGetShapeProjection(uint openXmlShapeId,
            out LegacyPptShapeProjection? projection);

        private const string ClassicAnimationExtensionUri =
            "{5BA743F1-2B69-4BB9-B2E0-4A418B7E7435}";
        private const string ClassicAnimationNamespace =
            "https://schemas.officeimo.net/powerpoint/2026/classic-animations";
        private readonly IReadOnlyDictionary<string, string> _slides;
        private readonly EditableImageFingerprintScope _editableImages;

        private LegacyPptProjectionFingerprint(string global,
            IReadOnlyDictionary<string, string> slides,
            EditableImageFingerprintScope editableImages) {
            Global = global;
            _slides = new ReadOnlyDictionary<string, string>(slides.ToDictionary(
                pair => pair.Key, pair => pair.Value, StringComparer.Ordinal));
            _editableImages = editableImages;
        }

        internal string Global { get; }

        internal static LegacyPptProjectionFingerprint Create(LegacyPptPackageView view,
            LegacyPptProjectionMap projectionMap) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(projectionMap);
            EditableImageFingerprintScope editableImages =
                EditableImageFingerprintScope.Create(view, projectionMap);
            var slides = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (LegacyPptPartView slidePart in LegacyPptPackageFingerprint.SlideParts(view)) {
                if (projectionMap.Slides.Any(slide => string.Equals(slide.SlidePartUri,
                        slidePart.Name, StringComparison.Ordinal))) {
                    slides.Add(slidePart.Name, CreateSlide(view,
                        slidePart, projectionMap, editableImages));
                }
            }
            if (slides.Count != projectionMap.Slides.Count) {
                throw new InvalidDataException("The projected slide fingerprint set is incomplete.");
            }
            return new LegacyPptProjectionFingerprint(CreateGlobal(view,
                projectionMap, editableImages), slides, editableImages);
        }

        internal bool Matches(LegacyPptPackageView view, LegacyPptProjectionMap projectionMap) {
            ArgumentNullException.ThrowIfNull(view);
            ArgumentNullException.ThrowIfNull(projectionMap);
            EditableImageFingerprintScope editableImages =
                EditableImageFingerprintScope.Create(view, projectionMap)
                    .Merge(_editableImages);
            if (!string.Equals(Global, CreateGlobal(view, projectionMap,
                    editableImages),
                    StringComparison.Ordinal)) return false;
            foreach (LegacyPptPartView slidePart in LegacyPptPackageFingerprint.SlideParts(view)) {
                if (_slides.TryGetValue(slidePart.Name, out string? expected)
                    && !string.Equals(expected, CreateSlide(view,
                        slidePart, projectionMap, editableImages),
                        StringComparison.Ordinal)) {
                    return false;
                }
            }
            return true;
        }

        private static string CreateGlobal(LegacyPptPackageView view,
            LegacyPptProjectionMap projectionMap,
            EditableImageFingerprintScope editableImages) {
            var materializedLayoutThemePartUris = new HashSet<string>(StringComparer.Ordinal);
            foreach (LegacyPptPartView master in LegacyPptPackageFingerprint.SlideMasterParts(view)) {
                foreach (LegacyPptPartView layout in LegacyPptPackageFingerprint.SlideLayoutParts(master)) {
                    if (!projectionMap.IsEditableProjectedLayoutThemePart(layout.Name)) continue;
                    LegacyPptPartView themeOverride = layout.ThemeOverride;
                    if (!themeOverride.IsNone) materializedLayoutThemePartUris.Add(themeOverride.Name);
                }
            }
            string presentationPartName = view.Presentation.Name;
            return LegacyPptPackageFingerprint.Create(view,
                (part, root) => {
                    if (string.Equals(part.Name, presentationPartName, StringComparison.OrdinalIgnoreCase)) {
                        NormalizeDocumentTopology(root);
                    }
                    if (part.HasContentType(LegacyPptPackageFingerprint.ExtendedPropertiesContentType)) {
                        NormalizeProjectedExtendedProperties(root);
                    }
                    if (part.IsSlideMaster
                        && projectionMap.TryGetMaster(part.Name,
                            out LegacyPptMasterProjection? masterProjection)
                        && masterProjection != null) {
                        NormalizeProjectedMaster(root, masterProjection);
                    }
                    if ((part.IsNotesMaster || part.IsHandoutMaster)
                        && projectionMap.TryGetSpecialMaster(part.Name,
                            out LegacyPptMasterProjection? specialProjection)
                        && specialProjection != null) {
                        NormalizeProjectedSpecialMaster(root, specialProjection);
                    }
                    if ((part.IsTheme || part.IsThemeOverride)
                        && projectionMap.IsProjectedMasterThemePart(part.Name)) {
                        NormalizeProjectedMasterTheme(root);
                    }
                    if (part.IsSlideLayout
                        && projectionMap.IsProjectedLayoutPart(part.Name)) {
                        NormalizeProjectedHeaderFooter(root);
                    }
                    if (part.IsSlideLayout
                        && projectionMap
                            .IsEditableProjectedOrdinaryLayoutPart(part.Name)) {
                        NormalizeProjectedOrdinaryLayout(root);
                    }
                    if (part.IsSlideLayout
                        && projectionMap.IsEditableProjectedLayoutBackgroundPart(part.Name)
                        && root.Is(LegacyPptXml.P, "sldLayout")) {
                        NormalizeProjectedBackground(root);
                    }
                    if (part.IsSlideLayout
                        && projectionMap.IsEditableProjectedLayoutThemePart(part.Name)) {
                        NormalizeProjectedLayoutTheme(root);
                    }
                    if (part.IsSlideLayout
                        && projectionMap.TryGetTitleMaster(part.Name,
                            out LegacyPptMasterProjection? titleProjection)
                        && titleProjection != null) {
                        NormalizeProjectedTitleMaster(root, titleProjection);
                    }
                },
                part => !(part.IsSlide || part.IsNotesSlide
                        || LegacyPptPackageFingerprint.IsSlideComments(part)
                        || LegacyPptPackageFingerprint.IsCommentAuthors(part)
                        || part.HasContentType(LegacyPptPackageFingerprint.CorePropertiesContentType)
                        || part.HasContentType(LegacyPptPackageFingerprint.CustomPropertiesContentType)
                        || LegacyPptPackageFingerprint.IsVbaProject(part))
                    && !projectionMap.IsProjectedOlePart(part.Name)
                    && !materializedLayoutThemePartUris.Contains(part.Name)
                    && !editableImages.IsExclusiveImagePart(part),
                (owner, relationship) => !(relationship.TargetPart.IsSlide
                        || LegacyPptPackageFingerprint.IsSlideComments(relationship.TargetPart)
                        || LegacyPptPackageFingerprint.IsCommentAuthors(relationship.TargetPart)
                        || LegacyPptPackageFingerprint.IsVbaProject(relationship.TargetPart))
                    && !editableImages.IsEditableImageRelationship(owner,
                        relationship)
                    && !(owner.IsSlideLayout
                        && relationship.TargetPart.IsThemeOverride
                        && projectionMap.IsEditableProjectedLayoutThemePart(
                            owner.Name)),
                includeReferenceRelationship: null,
                includeDataPartReferenceRelationship: null);
        }

        private static void NormalizeProjectedMaster(OfficeXmlNode root,
            LegacyPptMasterProjection projection) {
            if (!root.Is(LegacyPptXml.P, "sldMaster")) return;
            Clear(root.Child(LegacyPptXml.P, "clrMap"));
            NormalizeProjectedBackground(root);
            NormalizeProjectedShapes(root, projection.TryGetShape,
                normalizeInteractions: false);
            if (projection.CanEditTextStyles) {
                Remove(root.Child(LegacyPptXml.P, "txStyles"));
            }
        }

        private static void NormalizeProjectedMasterTheme(OfficeXmlNode root) {
            if (root.Is(LegacyPptXml.A, "theme") || root.Is(LegacyPptXml.A, "themeOverride")) {
                Clear(root);
            }
        }

        private static void NormalizeProjectedTitleMaster(OfficeXmlNode root,
            LegacyPptMasterProjection projection) {
            if (!root.Is(LegacyPptXml.P, "sldLayout")) return;
            root.RemoveAttribute("showMasterSp");
            Clear(root.Child(LegacyPptXml.P, "clrMapOvr"));
            NormalizeProjectedBackground(root);
            NormalizeProjectedShapes(root, projection.TryGetShape,
                normalizeInteractions: false);
        }

        private static void NormalizeProjectedSpecialMaster(OfficeXmlNode root,
            LegacyPptMasterProjection projection) {
            if (root.Is(LegacyPptXml.P, "notesMaster") || root.Is(LegacyPptXml.P, "handoutMaster")) {
                Clear(root.Child(LegacyPptXml.P, "clrMap"));
                NormalizeProjectedBackground(root);
            }
            NormalizeProjectedShapes(root, projection.TryGetShape,
                normalizeInteractions: false);
        }

        /// <summary>Removes the background of a slide-like root's common slide data.</summary>
        private static void NormalizeProjectedBackground(OfficeXmlNode root) =>
            Remove(root.Child(LegacyPptXml.P, "cSld").Child(LegacyPptXml.P, "bg"));

        private static string CreateSlide(LegacyPptPackageView view, LegacyPptPartView slidePart,
            LegacyPptProjectionMap projectionMap,
            EditableImageFingerprintScope editableImages) =>
            LegacyPptPackageFingerprint.Create(view,
            (part, root) => NormalizeProjectedSlide(root, slidePart.Name, projectionMap),
            part => part.Equals(slidePart)
                    || part.IsNotesSlide
                    && part.RelatedPart(LegacyPptPackageView.SlideRelationshipType).Equals(slidePart),
            (owner, relationship) => !LegacyPptPackageFingerprint.IsSlideComments(relationship.TargetPart)
                    && !relationship.TargetPart.IsSlide
                    && !editableImages.IsEditableImageRelationship(owner,
                        relationship),
            (owner, relationship) => !relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType),
            (owner, relationship) =>
                !IsWriterSupportedAudioDataRelationship(owner,
                    relationship));

        private static bool IsWriterSupportedAudioDataRelationship(
            LegacyPptPartView owner,
            LegacyPptRelationshipView relationship) {
            if (!owner.IsSlide) return false;
            OfficeXmlNode root = owner.Root;
            if (root.IsNone) return false;
            var referencedAudioIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode sound in root.SchemaDescendants(LegacyPptXml.P, "snd")) {
                AddIfPresent(referencedAudioIds, sound.Attribute(LegacyPptXml.R, "embed"));
            }
            foreach (OfficeXmlNode sound in root.SchemaDescendants(LegacyPptXml.A, "snd")) {
                AddIfPresent(referencedAudioIds, sound.Attribute(LegacyPptXml.R, "embed"));
            }
            foreach (OfficeXmlNode element in root.EveryDescendant()) {
                if (element.LocalName is not "snd" and not "sound") continue;
                AddIfPresent(referencedAudioIds, FirstAttributeNamed(element, "embed"));
            }
            if (referencedAudioIds.Count == 0) return false;
            LegacyPptRelationshipView[] dataRelationships = owner.Relationships
                .Where(candidate => LegacyPptPackageFingerprint.IsDataPartReference(candidate)
                    && candidate.TargetPart.Equals(relationship.TargetPart)).ToArray();
            return dataRelationships.Length > 0
                && dataRelationships.All(candidate =>
                    candidate.IsType(LegacyPptPackageView.AudioRelationshipType)
                    && referencedAudioIds.Contains(candidate.Id));
        }

        private static void AddIfPresent(ISet<string> set, string? value) {
            if (!string.IsNullOrEmpty(value)) set.Add(value!);
        }

        private static string? FirstAttributeNamed(OfficeXmlNode element, string localName) {
            if (element.Tree is not { } tree) return null;
            int count = tree.AttributeCountOf(element.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(element.Node, index, out int namespaceId, out string name, out string value);
                if (namespaceId == OfficeXmlNamespaces.Xmlns) continue;
                if (string.Equals(name, localName, StringComparison.Ordinal)) return value;
            }
            return null;
        }

        private sealed class EditableImageFingerprintScope {
            private readonly ISet<string> _relationshipKeys;
            private readonly ISet<string> _exclusiveImagePartUris;

            private EditableImageFingerprintScope(
                ISet<string> exclusiveImagePartUris,
                ISet<string> relationshipKeys) {
                _exclusiveImagePartUris = exclusiveImagePartUris;
                _relationshipKeys = relationshipKeys;
            }

            internal bool IsExclusiveImagePart(LegacyPptPartView part) =>
                _exclusiveImagePartUris.Contains(part.Name);

            internal bool IsEditableImageRelationship(LegacyPptPartView owner,
                LegacyPptRelationshipView relationship) => _relationshipKeys.Contains(
                    CreateRelationshipKey(owner, relationship.Id));

            internal EditableImageFingerprintScope Merge(
                EditableImageFingerprintScope other) {
                ArgumentNullException.ThrowIfNull(other);
                return new EditableImageFingerprintScope(
                    new HashSet<string>(_exclusiveImagePartUris.Concat(
                        other._exclusiveImagePartUris),
                        StringComparer.Ordinal),
                    new HashSet<string>(_relationshipKeys.Concat(
                        other._relationshipKeys),
                        StringComparer.Ordinal));
            }

            internal static EditableImageFingerprintScope Create(
                LegacyPptPackageView view,
                LegacyPptProjectionMap projectionMap) {
                var relationshipKeys = new HashSet<string>(
                    StringComparer.Ordinal);
                ISet<string> editableBackgroundOwners =
                    CollectEditableBackgroundOwners(view, projectionMap);
                var incoming = new Dictionary<string, List<bool>>(StringComparer.Ordinal);
                foreach (LegacyPptPartView owner in LegacyPptPackageFingerprint.CollectParts(view)) {
                    foreach (LegacyPptRelationshipView relationship in owner.Relationships) {
                        if (!LegacyPptPackageFingerprint.IsPartRelationship(relationship)
                            || !relationship.IsType(LegacyPptPackageView.ImageRelationshipType)) {
                            continue;
                        }
                        bool editableOnly = IsUsedOnlyByPictureBullets(owner,
                                relationship.Id)
                            || IsUsedOnlyByEditableBackgrounds(owner,
                                relationship.Id,
                                editableBackgroundOwners);
                        string image = relationship.TargetPart.Name;
                        if (!incoming.TryGetValue(image,
                                out List<bool>? usages)) {
                            usages = new List<bool>();
                            incoming.Add(image, usages);
                        }
                        usages.Add(editableOnly);
                        if (editableOnly) relationshipKeys.Add(
                            CreateRelationshipKey(owner,
                                relationship.Id));
                    }
                }
                ISet<string> exclusive = new HashSet<string>(
                    incoming.Where(pair => pair.Value.Count > 0
                            && pair.Value.All(value => value))
                        .Select(pair => pair.Key),
                    StringComparer.Ordinal);
                return new EditableImageFingerprintScope(exclusive,
                    relationshipKeys);
            }

            private static bool IsUsedOnlyByEditableBackgrounds(
                LegacyPptPartView owner, string relationshipId,
                ISet<string> editableBackgroundOwners) {
                if (!editableBackgroundOwners.Contains(owner.Name)) {
                    return false;
                }
                OfficeXmlNode[] references = BlipsReferencing(owner, relationshipId);
                return references.Length > 0 && references.All(blip =>
                    blip.HasAncestor(LegacyPptXml.P, "bg"));
            }

            private static ISet<string> CollectEditableBackgroundOwners(
                LegacyPptPackageView view,
                LegacyPptProjectionMap projectionMap) {
                var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (view.Presentation.IsNone) return result;
                foreach (LegacyPptPartView slidePart in LegacyPptPackageFingerprint.SlideParts(view)) {
                    if (!projectionMap.Slides.Any(slide => string.Equals(
                        slide.SlidePartUri, slidePart.Name,
                        StringComparison.Ordinal))) continue;
                    result.Add(slidePart.Name);
                    LegacyPptPartView notes = slidePart.NotesSlide;
                    if (!notes.IsNone) result.Add(notes.Name);
                }
                foreach (LegacyPptPartView masterPart in LegacyPptPackageFingerprint.SlideMasterParts(view)) {
                    if (projectionMap.TryGetMaster(masterPart.Name, out _)) {
                        result.Add(masterPart.Name);
                    }
                    foreach (LegacyPptPartView layoutPart in LegacyPptPackageFingerprint.SlideLayoutParts(masterPart)) {
                        if (projectionMap.TryGetTitleMaster(layoutPart.Name, out _)
                            || projectionMap
                                .IsEditableProjectedLayoutBackgroundPart(
                                    layoutPart.Name)) {
                            result.Add(layoutPart.Name);
                        }
                    }
                }
                LegacyPptPartView notesMaster = view.NotesMaster;
                if (!notesMaster.IsNone
                    && projectionMap.TryGetSpecialMaster(notesMaster.Name, out _)) {
                    result.Add(notesMaster.Name);
                }
                LegacyPptPartView handoutMaster = view.HandoutMaster;
                if (!handoutMaster.IsNone
                    && projectionMap.TryGetSpecialMaster(handoutMaster.Name, out _)) {
                    result.Add(handoutMaster.Name);
                }
                return result;
            }

            private static bool IsUsedOnlyByPictureBullets(
                LegacyPptPartView owner, string relationshipId) {
                OfficeXmlNode[] references = BlipsReferencing(owner, relationshipId);
                return references.Length > 0 && references.All(blip =>
                    blip.HasAncestor(LegacyPptXml.A, "buBlip"));
            }

            private static OfficeXmlNode[] BlipsReferencing(LegacyPptPartView owner, string relationshipId) {
                OfficeXmlNode root = owner.Root;
                if (root.IsNone) return Array.Empty<OfficeXmlNode>();
                return root.SchemaDescendants(LegacyPptXml.A, "blip")
                    .Where(blip => string.Equals(blip.Attribute(LegacyPptXml.R, "embed"),
                        relationshipId, StringComparison.Ordinal))
                    .ToArray();
            }

            private static string CreateRelationshipKey(LegacyPptPartView owner,
                string relationshipId) => owner.Name + "|" + relationshipId;
        }

        private static void NormalizeDocumentTopology(OfficeXmlNode root) {
            if (!root.Is(LegacyPptXml.P, "presentation")) return;
            OfficeXmlNode slideIds = root.Child(LegacyPptXml.P, "sldIdLst");
            if (!slideIds.IsNone) slideIds.RemoveChildren(LegacyPptXml.P, "sldId");
            root.RemoveChildren(LegacyPptXml.P, "custShowLst");
        }

        private static void NormalizeProjectedExtendedProperties(
            OfficeXmlNode root) {
            if (!root.Is(OfficeXmlNamespaces.ExtendedProperties, "Properties")) return;
            foreach (string name in new[] {
                         "Application", "TotalTime", "PresentationFormat", "Slides",
                         "Notes", "HiddenSlides", "Manager", "Company"
                     }) {
                Remove(root.Child(OfficeXmlNamespaces.ExtendedProperties, name));
            }
        }

        private static void NormalizeProjectedHeaderFooter(OfficeXmlNode root) {
            if (!root.Is(LegacyPptXml.P, "sldLayout")) return;
            root.RemoveChildren(LegacyPptXml.P, "hf");
            OfficeXmlNode tree = root.Path(LegacyPptXml.P, "cSld", "spTree");
            foreach (OfficeXmlNode shape in tree.Children(LegacyPptXml.P, "sp").ToList()) {
                string? type = shape.Path(LegacyPptXml.P, "nvSpPr", "nvPr", "ph").Attribute("type");
                if (type is not "dt" and not "ftr" and not "sldNum") continue;
                OfficeXmlNode text = shape.Child(LegacyPptXml.P, "txBody");
                if (!text.IsNone) text.RemoveChildren(LegacyPptXml.A, "p");
            }
        }

        private static void NormalizeProjectedOrdinaryLayout(
            OfficeXmlNode root) {
            if (!root.Is(LegacyPptXml.P, "sldLayout")) return;
            root.RemoveAttribute("type");
            OfficeXmlNode tree = root.Path(LegacyPptXml.P, "cSld", "spTree");
            if (!tree.IsNone) RemoveAllChildren(tree);
        }

        private static void NormalizeProjectedLayoutTheme(OfficeXmlNode root) {
            if (!root.Is(LegacyPptXml.P, "sldLayout")) return;
            Remove(root.Child(LegacyPptXml.P, "clrMapOvr"));
        }

        private static void NormalizeProjectedSlide(OfficeXmlNode root, string partName,
            LegacyPptProjectionMap projectionMap) {
            LegacyPptSlideProjection? slideProjection = projectionMap.Slides.FirstOrDefault(slide =>
                string.Equals(slide.SlidePartUri, partName, StringComparison.Ordinal));
            if (slideProjection == null) return;
            if (root.Is(LegacyPptXml.P, "notes") && slideProjection.Notes != null) {
                root.RemoveAttribute("showMasterSp");
                Clear(root.Child(LegacyPptXml.P, "clrMapOvr"));
                NormalizeProjectedBackground(root);
                OfficeXmlNode tree = root.Path(LegacyPptXml.P, "cSld", "spTree");
                foreach (OfficeXmlNode shape in tree.Children(LegacyPptXml.P, "sp").ToList()) {
                    OfficeXmlNode placeholder = shape.Path(LegacyPptXml.P, "nvSpPr", "nvPr", "ph");
                    OfficeXmlNode text = shape.Child(LegacyPptXml.P, "txBody");
                    if (placeholder.Attribute("type") == "body" && !text.IsNone) {
                        text.RemoveChildren(LegacyPptXml.A, "p");
                    }
                }
                return;
            }
            if (root.Is(LegacyPptXml.P, "sld")) {
                root.RemoveAttribute("show");
                root.RemoveAttribute("showMasterSp");
                Remove(root.Child(LegacyPptXml.P, "transition"));
                Clear(root.Child(LegacyPptXml.P, "clrMapOvr"));
                NormalizeProjectedBackground(root);
                OfficeXmlNode extensions = root.Child(LegacyPptXml.P, "extLst");
                OfficeXmlNode classicAnimations = default;
                foreach (OfficeXmlNode extension in extensions.Children(LegacyPptXml.P, "ext")) {
                    if (string.Equals(extension.Attribute("uri"),
                            ClassicAnimationExtensionUri, StringComparison.Ordinal)) {
                        classicAnimations = extension;
                        break;
                    }
                }
                if (!classicAnimations.IsNone
                    && HasOnlyTopLevelAnimationTargets(classicAnimations,
                        slideProjection)) {
                    Remove(root.Child(LegacyPptXml.P, "timing"));
                    classicAnimations.Remove();
                    if (extensions.FirstChild.IsNone) extensions.Remove();
                }
            }

            NormalizeProjectedShapes(root, slideProjection.TryGetShape,
                normalizeInteractions: true);
        }

        private static void NormalizeProjectedShapes(OfficeXmlNode root,
            TryGetShapeProjection tryGetShape, bool normalizeInteractions) {
            foreach (OfficeXmlNode shape in root.SchemaDescendants(LegacyPptXml.P, "sp")) {
                OfficeXmlNode nonVisual = shape.Child(LegacyPptXml.P, "nvSpPr");
                if (!TryGetProjection(nonVisual, tryGetShape,
                        out LegacyPptShapeProjection shapeProjection)) continue;
                Remove(nonVisual.Path(LegacyPptXml.P, "nvPr", "ph"));
                OfficeXmlNode drawing = nonVisual.Child(LegacyPptXml.P, "cNvPr");
                NormalizeShapeMetadata(drawing, shapeProjection);
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    RemoveShapeInteractions(drawing);
                }
                OfficeXmlNode properties = shape.Child(LegacyPptXml.P, "spPr");
                NormalizeOffsetAndExtents(properties.Child(LegacyPptXml.A, "xfrm"));
                NormalizeShapeFormatting(properties, shapeProjection);
                OfficeXmlNode text = shape.Child(LegacyPptXml.P, "txBody");
                if (!text.IsNone) {
                    bool retainsSecuritySensitiveTextInteractions =
                        !normalizeInteractions && text.EveryDescendant()
                            .Any(element => element.Is(LegacyPptXml.A, "hlinkClick")
                                || element.Is(LegacyPptXml.A, "hlinkHover")
                                || element.Is(LegacyPptXml.A, "hlinkMouseOver"));
                    if (!retainsSecuritySensitiveTextInteractions) {
                        text.RemoveChildren(LegacyPptXml.A, "p");
                    }
                    if (shapeProjection.CanEditTextFrame) {
                        Clear(text.Child(LegacyPptXml.A, "bodyPr"));
                    }
                }
            }
            foreach (OfficeXmlNode picture in root.SchemaDescendants(LegacyPptXml.P, "pic")) {
                OfficeXmlNode nonVisual = picture.Child(LegacyPptXml.P, "nvPicPr");
                if (!TryGetProjection(nonVisual, tryGetShape,
                        out LegacyPptShapeProjection shapeProjection)) continue;
                Remove(nonVisual.Path(LegacyPptXml.P, "nvPr", "ph"));
                OfficeXmlNode drawing = nonVisual.Child(LegacyPptXml.P, "cNvPr");
                NormalizeShapeMetadata(drawing, shapeProjection);
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    RemoveShapeInteractions(drawing);
                }
                OfficeXmlNode properties = picture.Child(LegacyPptXml.P, "spPr");
                NormalizeOffsetAndExtents(properties.Child(LegacyPptXml.A, "xfrm"));
                NormalizeShapeFormatting(properties, shapeProjection);
                if (shapeProjection.CanEditPictureFormatting) {
                    OfficeXmlNode pictureDrawing = nonVisual.Child(LegacyPptXml.P, "cNvPicPr");
                    if (!pictureDrawing.IsNone) {
                        pictureDrawing.RemoveAttribute("preferRelativeResize");
                        pictureDrawing.RemoveChildren(LegacyPptXml.A, "picLocks");
                    }
                    OfficeXmlNode fill = picture.Child(LegacyPptXml.P, "blipFill");
                    Remove(fill.Child(LegacyPptXml.A, "srcRect"));
                    OfficeXmlNode blip = fill.Child(LegacyPptXml.A, "blip");
                    if (!blip.IsNone) {
                        blip.RemoveChildren(LegacyPptXml.A, "lum");
                        blip.RemoveChildren(LegacyPptXml.A, "grayscl");
                        blip.RemoveChildren(LegacyPptXml.A, "biLevel");
                        blip.RemoveChildren(LegacyPptXml.A, "clrChange");
                        blip.RemoveChildren(LegacyPptXml.A, "clrRepl");
                    }
                }
            }
            foreach (OfficeXmlNode connection in root.SchemaDescendants(LegacyPptXml.P, "cxnSp")) {
                OfficeXmlNode nonVisual = connection.Child(LegacyPptXml.P, "nvCxnSpPr");
                if (!TryGetProjection(nonVisual, tryGetShape,
                        out LegacyPptShapeProjection shapeProjection)) continue;
                Remove(nonVisual.Path(LegacyPptXml.P, "nvPr", "ph"));
                OfficeXmlNode drawing = nonVisual.Child(LegacyPptXml.P, "cNvPr");
                NormalizeShapeMetadata(drawing, shapeProjection);
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    RemoveShapeInteractions(drawing);
                }
                OfficeXmlNode properties = connection.Child(LegacyPptXml.P, "spPr");
                NormalizeOffsetAndExtents(properties.Child(LegacyPptXml.A, "xfrm"));
                NormalizeShapeFormatting(properties, shapeProjection);
            }
            foreach (OfficeXmlNode group in root.SchemaDescendants(LegacyPptXml.P, "grpSp")) {
                OfficeXmlNode nonVisual = group.Child(LegacyPptXml.P, "nvGrpSpPr");
                if (!TryGetProjection(nonVisual, tryGetShape,
                        out LegacyPptShapeProjection shapeProjection)) continue;
                OfficeXmlNode drawing = nonVisual.Child(LegacyPptXml.P, "cNvPr");
                NormalizeShapeMetadata(drawing, shapeProjection);
                if (normalizeInteractions && shapeProjection.CanEditInteractions) {
                    RemoveShapeInteractions(drawing);
                }
                OfficeXmlNode transform = group.Path(LegacyPptXml.P, "grpSpPr").Child(LegacyPptXml.A, "xfrm");
                NormalizeOffsetAndExtents(transform);
                if (shapeProjection.CanEditShapeTransform && !transform.IsNone) {
                    RemoveRotationAndFlips(transform);
                }
                if (shapeProjection.CanEditGroupCoordinate && !transform.IsNone) {
                    ZeroPoint(transform.Child(LegacyPptXml.A, "chOff"));
                    ZeroSize(transform.Child(LegacyPptXml.A, "chExt"));
                }
            }
            foreach (OfficeXmlNode frame in root.SchemaDescendants(LegacyPptXml.P, "graphicFrame")) {
                OfficeXmlNode nonVisual = frame.Child(LegacyPptXml.P, "nvGraphicFramePr");
                if (!TryGetProjection(nonVisual, tryGetShape,
                        out LegacyPptShapeProjection shapeProjection)) continue;
                NormalizeShapeMetadata(nonVisual.Child(LegacyPptXml.P, "cNvPr"), shapeProjection);
                if (shapeProjection.OleObject == null) continue;
                OfficeXmlNode transform = frame.Child(LegacyPptXml.P, "xfrm");
                NormalizeOffsetAndExtents(transform);
                if (shapeProjection.CanEditShapeTransform && !transform.IsNone) {
                    RemoveRotationAndFlips(transform);
                }
                OfficeXmlNode ole = frame.Child(LegacyPptXml.A, "graphic")
                    .Child(LegacyPptXml.A, "graphicData").Child(LegacyPptXml.P, "oleObj");
                if (ole.IsNone) continue;
                ole.RemoveAttribute("progId");
                ole.RemoveAttribute("showAsIcon");
                OfficeXmlNode embed = ole.Child(LegacyPptXml.P, "embed");
                if (!embed.IsNone) embed.RemoveAttribute("followColorScheme");
            }
        }

        private static bool TryGetProjection(OfficeXmlNode nonVisual, TryGetShapeProjection tryGetShape,
            out LegacyPptShapeProjection projection) {
            projection = null!;
            uint? shapeId = nonVisual.Child(LegacyPptXml.P, "cNvPr").UInt("id");
            if (!shapeId.HasValue || !tryGetShape(shapeId.Value,
                    out LegacyPptShapeProjection? found) || found == null) return false;
            projection = found;
            return true;
        }

        private static void RemoveShapeInteractions(OfficeXmlNode drawing) {
            if (drawing.IsNone) return;
            drawing.RemoveChildren(LegacyPptXml.A, "hlinkClick");
            drawing.RemoveChildren(LegacyPptXml.A, "hlinkHover");
        }

        private static void NormalizeOffsetAndExtents(OfficeXmlNode transform) {
            if (transform.IsNone) return;
            ZeroPoint(transform.Child(LegacyPptXml.A, "off"));
            ZeroSize(transform.Child(LegacyPptXml.A, "ext"));
        }

        private static void ZeroPoint(OfficeXmlNode point) {
            if (point.IsNone) return;
            point.SetAttribute("x", "0");
            point.SetAttribute("y", "0");
        }

        private static void ZeroSize(OfficeXmlNode size) {
            if (size.IsNone) return;
            size.SetAttribute("cx", "0");
            size.SetAttribute("cy", "0");
        }

        private static void RemoveRotationAndFlips(OfficeXmlNode transform) {
            transform.RemoveAttribute("rot");
            transform.RemoveAttribute("flipH");
            transform.RemoveAttribute("flipV");
        }

        private static void NormalizeShapeMetadata(
            OfficeXmlNode properties,
            LegacyPptShapeProjection projection) {
            if (properties.IsNone) return;
            if (projection.CanEditShapeMetadata) {
                properties.SetAttribute("name", string.Empty);
                properties.RemoveAttribute("descr");
            }
            if (projection.CanEditShapeVisibility) {
                properties.RemoveAttribute("hidden");
            }
        }

        private static void NormalizeShapeFormatting(
            OfficeXmlNode properties,
            LegacyPptShapeProjection projection) {
            if (properties.IsNone) return;
            OfficeXmlNode transform = properties.Child(LegacyPptXml.A, "xfrm");
            if (projection.CanEditShapeTransform && !transform.IsNone) {
                RemoveRotationAndFlips(transform);
            }
            if (projection.CanEditShapeGeometry) {
                OfficeXmlNode adjustments = properties.Child(LegacyPptXml.A, "prstGeom").Child(LegacyPptXml.A, "avLst");
                if (!adjustments.IsNone) adjustments.RemoveChildren(LegacyPptXml.A, "gd");
            }
            if (!projection.CanEditShapeVisualStyle) return;
            properties.RemoveChildren(LegacyPptXml.A, "noFill");
            properties.RemoveChildren(LegacyPptXml.A, "solidFill");
            properties.RemoveChildren(LegacyPptXml.A, "ln");
            properties.RemoveChildren(LegacyPptXml.A, "effectLst");
        }

        private static bool HasOnlyTopLevelAnimationTargets(
            OfficeXmlNode extension,
            LegacyPptSlideProjection slideProjection) {
            OfficeXmlNode[] animations = extension.EveryDescendant()
                .Where(element => element.Is(ClassicAnimationNamespace, "animation"))
                .ToArray();
            return animations.Length > 0 && animations.All(animation =>
                uint.TryParse(FirstAttributeNamed(animation, "shapeId"),
                    NumberStyles.Integer, CultureInfo.InvariantCulture,
                    out uint shapeId)
                && slideProjection.TryGetShape(shapeId, out _));
        }

        private static void Remove(OfficeXmlNode element) {
            if (!element.IsNone) element.Remove();
        }

        private static void RemoveAllChildren(OfficeXmlNode element) {
            if (element.Tree is not { } tree) return;
            while (tree.FirstChildOf(element.Node) != OfficeXmlTree.NoNode) {
                tree.Remove(tree.FirstChildOf(element.Node));
            }
        }

        /// <summary>Removes every attribute and child of an element, as the SDK's clear-all calls did.</summary>
        internal static void Clear(OfficeXmlNode element) {
            if (element.Tree is not { } tree) return;
            RemoveAllChildren(element);
            for (int index = tree.AttributeCountOf(element.Node) - 1; index >= 0; index--) {
                tree.GetAttributeAt(element.Node, index, out int namespaceId, out string localName, out _);
                if (namespaceId == OfficeXmlNamespaces.Xmlns
                    || namespaceId == OfficeXmlNamespaces.None && string.Equals(localName, "xmlns", StringComparison.Ordinal)) {
                    continue;
                }
                tree.RemoveAttribute(element.Node, namespaceId, localName);
            }
        }
    }

    /// <summary>
    /// A deterministic fingerprint over a presentation package read through a <see cref="LegacyPptPackageView"/>: parts,
    /// their canonical markup or bytes, and their relationships, independent of how the package was serialized. Part
    /// relationships, data-part (media) references and external or hyperlink references are told apart as the Open XML
    /// SDK models them.
    /// </summary>
    internal static class LegacyPptPackageFingerprint {
        internal const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        internal const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";
        internal const string CustomPropertiesContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        private const string SlideCommentsContentType = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
        private const string CommentAuthorsContentType = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
        private const string VbaProjectContentType = "application/vnd.ms-office.vbaProject";
        private const int MaximumPartCount = 100000;

        internal static bool IsSlideComments(LegacyPptPartView part) => part.HasContentType(SlideCommentsContentType);

        internal static bool IsCommentAuthors(LegacyPptPartView part) => part.HasContentType(CommentAuthorsContentType);

        internal static bool IsVbaProject(LegacyPptPartView part) => part.HasContentType(VbaProjectContentType);

        /// <summary>Whether a relationship is an external or hyperlink reference rather than a part edge.</summary>
        internal static bool IsReference(LegacyPptRelationshipView relationship) =>
            relationship.IsExternal || relationship.TargetPart.IsNone
            || relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType);

        /// <summary>Whether a relationship refers to a shared media data part (audio, video or media).</summary>
        internal static bool IsDataPartReference(LegacyPptRelationshipView relationship) =>
            !IsReference(relationship)
            && (relationship.IsType(LegacyPptPackageView.AudioRelationshipType)
                || relationship.IsType(LegacyPptPackageView.VideoRelationshipType)
                || relationship.IsType(LegacyPptPackageView.MediaRelationshipType));

        /// <summary>Whether a relationship is a part edge: internal, to an existing part, and not a data-part reference.</summary>
        internal static bool IsPartRelationship(LegacyPptRelationshipView relationship) =>
            !IsReference(relationship) && !IsDataPartReference(relationship);

        /// <summary>Every part reachable from the package through part edges, each once.</summary>
        internal static IReadOnlyList<LegacyPptPartView> CollectParts(LegacyPptPackageView view) {
            var parts = new List<LegacyPptPartView>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<LegacyPptPartView>();
            foreach (LegacyPptRelationshipView relationship in view.PackageRelationships) {
                if (IsPartRelationship(relationship)) pending.Push(relationship.TargetPart);
            }
            while (pending.Count > 0) {
                LegacyPptPartView part = pending.Pop();
                if (!seen.Add(part.Name)) continue;
                if (seen.Count > MaximumPartCount) {
                    throw new IOException("The presentation part count exceeds the fingerprint limit.");
                }
                parts.Add(part);
                foreach (LegacyPptRelationshipView relationship in part.Relationships) {
                    if (IsPartRelationship(relationship)) pending.Push(relationship.TargetPart);
                }
            }
            return parts;
        }

        /// <summary>The presentation's slide parts in relationship order, as the SDK lists them.</summary>
        internal static List<LegacyPptPartView> SlideParts(LegacyPptPackageView view) =>
            view.Presentation.RelatedParts(LegacyPptPackageView.SlideRelationshipType);

        /// <summary>The presentation's slide master parts in relationship order, as the SDK lists them.</summary>
        internal static List<LegacyPptPartView> SlideMasterParts(LegacyPptPackageView view) =>
            view.Presentation.RelatedParts(LegacyPptPackageView.SlideMasterRelationshipType);

        /// <summary>A master's slide layout parts in relationship order, as the SDK lists them.</summary>
        internal static List<LegacyPptPartView> SlideLayoutParts(LegacyPptPartView master) =>
            master.RelatedParts(LegacyPptPackageView.SlideLayoutRelationshipType);

        /// <summary>A fingerprint of the whole package, every part and relationship included.</summary>
        internal static string Create(LegacyPptPackageView view) => Create(view, null, null, null, null, null);

        internal static string Create(LegacyPptPackageView view,
            Action<LegacyPptPartView, OfficeXmlNode>? normalizeRoot,
            Func<LegacyPptPartView, bool>? includePart,
            Func<LegacyPptPartView, LegacyPptRelationshipView, bool>? includeRelationship,
            Func<LegacyPptPartView, LegacyPptRelationshipView, bool>? includeReferenceRelationship,
            Func<LegacyPptPartView, LegacyPptRelationshipView, bool>? includeDataPartReferenceRelationship) {
            ArgumentNullException.ThrowIfNull(view);
            using IncrementalHash content = LegacyPptProjectionDigest.CreateBuilder();
            foreach (LegacyPptPartView part in CollectParts(view).OrderBy(item => item.Name, StringComparer.Ordinal)) {
                if (includePart != null && !includePart(part)) continue;
                LegacyPptProjectionDigest.Append(content, part.Name);
                LegacyPptProjectionDigest.Append(content, part.ContentType);
                AppendContent(content, part, normalizeRoot);

                foreach (LegacyPptRelationshipView relationship in part.Relationships
                             .Where(IsPartRelationship).OrderBy(item => item.Id, StringComparer.Ordinal)) {
                    if (includeRelationship != null && !includeRelationship(part, relationship)) continue;
                    LegacyPptProjectionDigest.Append(content, relationship.Id);
                    LegacyPptProjectionDigest.Append(content, relationship.TargetPart.Name);
                }
                foreach (LegacyPptRelationshipView relationship in part.Relationships
                             .Where(IsDataPartReference).OrderBy(item => item.Id, StringComparer.Ordinal)) {
                    if (includeDataPartReferenceRelationship != null
                        && !includeDataPartReferenceRelationship(part, relationship)) continue;
                    LegacyPptPartView dataPart = relationship.TargetPart;
                    LegacyPptProjectionDigest.Append(content, relationship.Id);
                    LegacyPptProjectionDigest.Append(content, OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type));
                    LegacyPptProjectionDigest.Append(content, dataPart.Name);
                    LegacyPptProjectionDigest.Append(content, dataPart.ContentType);
                    AppendBytes(content, dataPart);
                }
                foreach (LegacyPptRelationshipView relationship in part.Relationships
                             .Where(IsReference).OrderBy(item => item.Id, StringComparer.Ordinal)) {
                    if (includeReferenceRelationship != null
                        && !includeReferenceRelationship(part, relationship)) continue;
                    LegacyPptProjectionDigest.Append(content, relationship.Id);
                    LegacyPptProjectionDigest.Append(content, OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type));
                    LegacyPptProjectionDigest.Append(content, relationship.Target);
                }
            }
            return LegacyPptProjectionDigest.Finish(content);
        }

        private static void AppendContent(IncrementalHash content, LegacyPptPartView part,
            Action<LegacyPptPartView, OfficeXmlNode>? normalizeRoot) {
            OfficeXmlNode root;
            try {
                root = part.Root;
            } catch (Exception exception) when (exception is System.Xml.XmlException
                                                    or InvalidDataException or IOException) {
                LegacyPptProjectionDigest.Append(content, "unreadable");
                return;
            }
            if (root.IsNone) {
                AppendBytes(content, part);
                return;
            }
            if (normalizeRoot != null) {
                OfficeXmlNode normalized = root.Clone();
                normalizeRoot(part, normalized);
                root = normalized;
            }
            var markup = new StringBuilder();
            LegacyPptXml.AppendCanonical(root, markup);
            LegacyPptProjectionDigest.Append(content, "xml");
            LegacyPptProjectionDigest.Append(content, markup.ToString());
        }

        private static void AppendBytes(IncrementalHash content, LegacyPptPartView part) {
            byte[] bytes;
            try {
                bytes = part.GetBytes();
            } catch (Exception exception) when (exception is InvalidDataException or IOException) {
                LegacyPptProjectionDigest.Append(content, "unreadable");
                return;
            }
            LegacyPptProjectionDigest.Append(content, bytes.Length.ToString(CultureInfo.InvariantCulture));
            content.AppendData(bytes);
        }
    }
}
