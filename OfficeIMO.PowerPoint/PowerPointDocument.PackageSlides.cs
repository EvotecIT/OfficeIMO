using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Slides of a presentation whose SDK package is not open. Each slide part has one <see cref="PowerPointSlide"/>, which
    /// works on the document's package; when something opens the SDK, the same objects bind to their SDK parts, so slides
    /// a caller holds keep working.
    /// </summary>
    public sealed partial class PowerPointDocument {
        // Slides by part identity, and the slide list in order, read again after the list changes.
        private Dictionary<string, PowerPointSlide>? _packageSlides;
        private List<PowerPointSlide>? _packageSlideOrder;

        /// <summary>
        /// The package that holds what is read and edited without the SDK, opened on first use; null when the document
        /// can't be read that way. <paramref name="edit"/> also requires that it can be edited.
        /// </summary>
        private OfficeOpcEditablePackage? GetPackage(bool edit) {
            if (_disposed || !_sdkDeferred || _sdkDocument != null || edit && _accessMode != DocumentAccessMode.ReadWrite
                || _deferredOpenSettings?.MaxCharactersInPart > 0
                || _packageStream is not { CanRead: true, CanSeek: true } stream) {
                return null;
            }
            if (_pendingPackage == null) {
                stream.Position = 0;
                _pendingPackage = OfficeOpcEditablePackage.Open(stream, leaveOpen: true);
                _pendingSource = stream;
            }
            return _pendingPackage;
        }

        /// <summary>Whether slides and shapes can be edited in the package without the SDK.</summary>
        internal bool CanEditPackage => GetPackage(edit: true) != null;

        /// <summary>Records an edit made in the package, so the next save or SDK open writes it.</summary>
        internal void MarkPackageEdited() => _pendingEdits = true;

        /// <summary>A part's tree in the package read without the SDK, or null when the package can't be read that way.</summary>
        internal OfficeXmlTree? GetPackageXml(string partName) {
            if (GetPackage(edit: false) is not { } package || !package.PartExists(partName)) return null;
            try {
                return package.GetXml(partName);
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return null;
            }
        }

        /// <summary>
        /// A slide's notes tree when the model would read the notes without adding anything: the notes slide links back to
        /// the slide and to the presentation's notes master, which has a theme and is listed in the presentation. Null
        /// otherwise, and the model reads or creates them through the SDK.
        /// </summary>
        internal OfficeXmlTree? GetPackageNotes(string slidePart) {
            if (GetPackage(edit: false) is not { } package) return null;
            try {
                string? notes = FirstRelated(package, slidePart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesSlide", out _);
                PackageSlideList? slides = notes == null ? null : PackageSlideList.Read(package);
                if (notes == null || slides == null) return null;
                string? master = FirstRelated(package, slides.Main, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesMaster", out string? masterRelationship);
                if (master == null || masterRelationship == null
                    || FirstRelated(package, master, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "theme", out _) == null
                    || !Relates(package, notes, slidePart) || !Relates(package, notes, master)) {
                    return null;
                }
                OfficeXmlTree presentation = slides.Presentation;
                int list = presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "notesMasterIdLst");
                bool listed = false;
                for (int entry = list == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : presentation.FindChild(list, OfficeXmlNamespaces.PresentationML, "notesMasterId");
                     entry != OfficeXmlTree.NoNode && !listed; entry = presentation.FindNextSibling(entry, OfficeXmlNamespaces.PresentationML, "notesMasterId")) {
                    listed = presentation.GetAttribute(entry, OfficeXmlNamespaces.Relationships, "id") == masterRelationship;
                }
                if (!listed || !package.GetXml(master).IsElement(package.GetXml(master).Root, OfficeXmlNamespaces.PresentationML, "notesMaster")) return null;
                OfficeXmlTree tree = package.GetXml(notes);
                return tree.IsElement(tree.Root, OfficeXmlNamespaces.PresentationML, "notes") ? tree : null;
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return null;
            }
        }

        // The first internal relationship of a type, as the SDK's typed part properties find it.
        private static string? FirstRelated(OfficeOpcEditablePackage package, string source, string type, out string? id) {
            id = null;
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source);
            if (!package.PartExists(relationshipsPart)) return null;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                if (OfficeXmlNamespaces.NormalizeRelationshipType(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty) != type) continue;
                string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (target == null || !package.PartExists(target)) return null;
                id = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                return target;
            }
            return null;
        }

        // Whether a part has an internal relationship to another part.
        private static bool Relates(OfficeOpcEditablePackage package, string source, string target) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source);
            if (!package.PartExists(relationshipsPart)) return false;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? resolved = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (resolved != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(resolved), PartIdentity(target))) return true;
            }
            return false;
        }

        /// <summary>The slides in order, read from the package, or null when the SDK has to read them.</summary>
        private List<PowerPointSlide>? GetPackageSlides() {
            if (_packageSlideOrder != null && _sdkDocument == null) return _packageSlideOrder;
            if (GetPackage(edit: false) is not { } package) return null;
            PackageSlideList? list;
            try {
                list = PackageSlideList.Read(package);
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return null;
            }
            if (list == null) return null;
            _packageSlides ??= new Dictionary<string, PowerPointSlide>(OfficeAsciiIgnoreCaseComparer.Instance);
            var order = new List<PowerPointSlide>(list.Entries.Count);
            foreach ((_, _, string part) in list.Entries) {
                string identity = PartIdentity(part);
                if (!_packageSlides.TryGetValue(identity, out PowerPointSlide? slide)) {
                    slide = new PowerPointSlide(this, part);
                    _packageSlides.Add(identity, slide);
                }
                order.Add(slide);
            }
            return _packageSlideOrder = order;
        }

        /// <summary>Forgets the slide order after the slide list changed; <paramref name="removedPart"/> left the package.</summary>
        private void PackageSlidesChanged(string? removedPart = null) {
            _packageSlideOrder = null;
            if (removedPart != null) _packageSlides?.Remove(PartIdentity(removedPart));
        }

        /// <summary>Opens the SDK for a package-backed slide and returns the part it is bound to.</summary>
        internal SlidePart BindSdkSlide(PowerPointSlide slide) {
            ThrowIfDisposed();
            EnsureSdk();
            // Opening the SDK binds every slide still listed; one that is not was removed.
            return slide.IsPackageBacked
                ? throw new InvalidOperationException("The slide is no longer part of the presentation.")
                : slide.SlidePart;
        }

        /// <summary>The package-backed slide for an SDK slide part, which the SDK load binds instead of creating a new one.</summary>
        private PowerPointSlide? TakePackageSlide(SlidePart slidePart) {
            if (_packageSlides == null) return null;
            string identity = PartIdentity(slidePart.Uri.OriginalString);
            if (!_packageSlides.TryGetValue(identity, out PowerPointSlide? slide)) return null;
            slide.BindSdk(slidePart);
            return slide;
        }
    }
}
