using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>An object that works on the package and binds to its SDK element once the SDK opens.</summary>
    internal interface IPackageSdkBinding {
        void BindSdk();
    }

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

        // Whether notes were added without the SDK, so a save counts notes the loaded package had no content type for.
        private bool _packageNotesAdded;

        // Objects below the shapes (paragraphs) that work on the package and bind to the SDK when it opens.
        private List<WeakReference<IPackageSdkBinding>>? _packageBindings;

        internal void RegisterPackageBinding(IPackageSdkBinding binding) =>
            (_packageBindings ??= new List<WeakReference<IPackageSdkBinding>>()).Add(new WeakReference<IPackageSdkBinding>(binding));

        // Binds every object still alive while the slides the SDK opened are as the package wrote them.
        private void BindPackageObjects() {
            List<WeakReference<IPackageSdkBinding>>? bindings = _packageBindings;
            _packageBindings = null;
            if (bindings == null) return;
            foreach (WeakReference<IPackageSdkBinding> reference in bindings) {
                if (reference.TryGetTarget(out IPackageSdkBinding? binding)) binding.BindSdk();
            }
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

        /// <summary>
        /// Adds a notes slide to a slide that has none, as the notes model creates it on first use: an empty body
        /// placeholder, a relationship back to the slide and one to the presentation's notes master. Null, having changed
        /// nothing, when the slide has notes or the presentation lacks a notes master the model would only use as it is.
        /// </summary>
        internal OfficeXmlTree? TryCreatePackageNotes(string slidePart) {
            if (GetPackage(edit: true) is not { } package) return null;
            try {
                // Any notes relationship, even one whose part is missing, is the model's to handle.
                if (HasRelationshipOfType(package, slidePart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesSlide")) return null;
                // Notes go in a folder beside the slide's folder, which a slide at the package root doesn't have.
                string slideDirectory = slidePart.Substring(0, slidePart.LastIndexOf('/'));
                if (slideDirectory.LastIndexOf('/') < 0) return null;
                if (PackageSlideList.Read(package) is not { } slides) return null;
                string? master = FirstRelated(package, slides.Main, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesMaster", out string? masterRelationship);
                if (master == null || masterRelationship == null
                    || FirstRelated(package, master, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "theme", out _) == null
                    || !package.GetXml(master).IsElement(package.GetXml(master).Root, OfficeXmlNamespaces.PresentationML, "notesMaster")) {
                    return null;
                }
                OfficeXmlTree presentation = slides.Presentation;
                int list = presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "notesMasterIdLst");
                bool listed = false;
                for (int entry = list == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : presentation.FindChild(list, OfficeXmlNamespaces.PresentationML, "notesMasterId");
                     entry != OfficeXmlTree.NoNode && !listed; entry = presentation.FindNextSibling(entry, OfficeXmlNamespaces.PresentationML, "notesMasterId")) {
                    listed = presentation.GetAttribute(entry, OfficeXmlNamespaces.Relationships, "id") == masterRelationship;
                }
                if (!listed) return null;

                // The slide's first free rIdN, as the notes model picks it.
                var used = new HashSet<string>(StringComparer.Ordinal);
                string slideRelationships = OfficeOpcPartNames.GetRelationshipsPartName(slidePart);
                if (package.PartExists(slideRelationships)) {
                    OfficeXmlTree relationships = package.GetXml(slideRelationships);
                    for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                         node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                        if (relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id") is { } existing) used.Add(existing);
                    }
                }
                int number = 1;
                string notesRelationship;
                do {
                    notesRelationship = "rId" + number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    number++;
                } while (used.Contains(notesRelationship));

                const string notesContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
                string slideFolder = slidePart.Substring(0, slidePart.LastIndexOf('/'));
                string notesPart = NextPartName(package, notesContentType, slideFolder.Substring(0, slideFolder.LastIndexOf('/') + 1) + "notesSlides/notesSlide", ".xml");
                OfficeXmlTree notes = package.CreateXmlPart(notesPart, notesContentType, OfficeXmlNamespaces.PresentationML, "notes", 32);
                WriteBlankNotes(notes);
                package.AddRelationship(slidePart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesSlide",
                    OfficeOpcPartNames.GetRelativeTarget(slidePart, notesPart), id: notesRelationship);
                package.AddRelationship(notesPart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slide",
                    OfficeOpcPartNames.GetRelativeTarget(notesPart, slidePart), id: "R" + Guid.NewGuid().ToString("N").Substring(0, 16));
                package.AddRelationship(notesPart, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesMaster",
                    OfficeOpcPartNames.GetRelativeTarget(notesPart, master), id: "R" + Guid.NewGuid().ToString("N").Substring(0, 16));
                _pendingRelationships = null;
                _pendingEdits = true;
                _packageNotesAdded = true;
                return notes;
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return null;
            }
        }

        // The notes slide the notes model writes: a group named for the notes, and the body placeholder with one empty run.
        private static void WriteBlankNotes(OfficeXmlTree notes) {
            int root = notes.Root;
            notes.DeclareNamespace(root, "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            notes.DeclareNamespace(root, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            int tree = Append(notes, Append(notes, root, OfficeXmlNamespaces.PresentationML, "cSld"), OfficeXmlNamespaces.PresentationML, "spTree");
            int group = Append(notes, tree, OfficeXmlNamespaces.PresentationML, "nvGrpSpPr");
            int groupDrawing = Append(notes, group, OfficeXmlNamespaces.PresentationML, "cNvPr");
            notes.SetAttribute(groupDrawing, OfficeXmlNamespaces.None, "id", "1");
            notes.SetAttribute(groupDrawing, OfficeXmlNamespaces.None, "name", "Notes Group Shape");
            Append(notes, group, OfficeXmlNamespaces.PresentationML, "cNvGrpSpPr");
            Append(notes, group, OfficeXmlNamespaces.PresentationML, "nvPr");
            int transform = Append(notes, Append(notes, tree, OfficeXmlNamespaces.PresentationML, "grpSpPr"), OfficeXmlNamespaces.DrawingML, "xfrm");
            AppendPoint(notes, transform, "off", "x", "y");
            AppendPoint(notes, transform, "ext", "cx", "cy");
            AppendPoint(notes, transform, "chOff", "x", "y");
            AppendPoint(notes, transform, "chExt", "cx", "cy");
            int shape = Append(notes, tree, OfficeXmlNamespaces.PresentationML, "sp");
            int nonVisual = Append(notes, shape, OfficeXmlNamespaces.PresentationML, "nvSpPr");
            int drawing = Append(notes, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPr");
            notes.SetAttribute(drawing, OfficeXmlNamespaces.None, "id", "2");
            notes.SetAttribute(drawing, OfficeXmlNamespaces.None, "name", "Notes Placeholder 1");
            notes.SetAttribute(Append(notes, Append(notes, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvSpPr"), OfficeXmlNamespaces.DrawingML, "spLocks"),
                OfficeXmlNamespaces.None, "noGrp", "1");
            int placeholder = Append(notes, Append(notes, nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr"), OfficeXmlNamespaces.PresentationML, "ph");
            notes.SetAttribute(placeholder, OfficeXmlNamespaces.None, "type", "body");
            notes.SetAttribute(placeholder, OfficeXmlNamespaces.None, "idx", "1");
            Append(notes, shape, OfficeXmlNamespaces.PresentationML, "spPr");
            int body = Append(notes, shape, OfficeXmlNamespaces.PresentationML, "txBody");
            Append(notes, body, OfficeXmlNamespaces.DrawingML, "bodyPr");
            Append(notes, body, OfficeXmlNamespaces.DrawingML, "lstStyle");
            int paragraph = Append(notes, body, OfficeXmlNamespaces.DrawingML, "p");
            int run = Append(notes, paragraph, OfficeXmlNamespaces.DrawingML, "r");
            notes.SetAttribute(Append(notes, run, OfficeXmlNamespaces.DrawingML, "rPr"), OfficeXmlNamespaces.None, "lang", "en-US");
            Append(notes, run, OfficeXmlNamespaces.DrawingML, "t");
            notes.SetAttribute(Append(notes, paragraph, OfficeXmlNamespaces.DrawingML, "endParaRPr"), OfficeXmlNamespaces.None, "lang", "en-US");
            Append(notes, Append(notes, root, OfficeXmlNamespaces.PresentationML, "clrMapOvr"), OfficeXmlNamespaces.DrawingML, "masterClrMapping");
        }

        // The first internal relationship of a type, as the SDK's typed part properties find it.
        private static string? FirstRelated(OfficeOpcEditablePackage package, string? source, string type, out string? id) {
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

        private static bool HasRelationshipOfType(OfficeOpcEditablePackage package, string source, string type) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source);
            if (!package.PartExists(relationshipsPart)) return false;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (OfficeXmlNamespaces.NormalizeRelationshipType(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty) == type) return true;
            }
            return false;
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

        /// <summary>The presentation part for a slide size that has to read or edit it through the SDK.</summary>
        internal DocumentFormat.OpenXml.Packaging.PresentationPart SdkPresentationPartForSize => SdkPresentationPart;

        /// <summary>The attributes of the presentation's first p:sldSz in the package, or null when it has none there.</summary>
        internal (string? Width, string? Height, string? Type)? GetPackageSlideSize() {
            if (GetPackage(edit: false) is not { } package) return null;
            try {
                if (PackageSlideList.Read(package) is not { } slides) return null;
                OfficeXmlTree presentation = slides.Presentation;
                int size = presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "sldSz");
                if (size == OfficeXmlTree.NoNode) return null;
                return (presentation.GetAttribute(size, OfficeXmlNamespaces.None, "cx"), presentation.GetAttribute(size, OfficeXmlNamespaces.None, "cy"),
                    presentation.GetAttribute(size, OfficeXmlNamespaces.None, "type"));
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return null;
            }
        }

        /// <summary>
        /// Sets or removes an attribute of the presentation's p:sldSz in the package; false when the presentation has none
        /// there, and the model adds one through the SDK.
        /// </summary>
        internal bool TrySetPackageSlideSize(string attribute, string? value) {
            if (GetPackage(edit: true) is not { } package) return false;
            try {
                if (PackageSlideList.Read(package) is not { } slides) return false;
                OfficeXmlTree presentation = slides.Presentation;
                int size = presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "sldSz");
                if (size == OfficeXmlTree.NoNode) return false;
                if (value == null) presentation.RemoveAttribute(size, OfficeXmlNamespaces.None, attribute);
                else presentation.SetAttribute(size, OfficeXmlNamespaces.None, attribute, value);
                _pendingEdits = true;
                return true;
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return false;
            }
        }

        /// <summary>A slide's entry in the presentation's slide list in the package, or null when it can't be read there.</summary>
        internal (OfficeXmlTree Tree, int Node)? GetPackageSlideEntry(string slidePart) {
            if (GetPackage(edit: false) is not { } package) return null;
            try {
                if (PackageSlideList.Read(package) is not { } slides) return null;
                string identity = PartIdentity(slidePart);
                foreach ((int node, _, string part) in slides.Entries) {
                    if (OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(part), identity)) return (slides.Presentation, node);
                }
                return null;
            } catch (Exception exception) when (exception is System.Xml.XmlException or InvalidDataException) {
                return null;
            }
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
