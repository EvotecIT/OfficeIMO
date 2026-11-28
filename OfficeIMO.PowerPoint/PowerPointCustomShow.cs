using System.Collections.ObjectModel;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>Represents an editable named slide sequence in a PowerPoint presentation.</summary>
    public sealed class PowerPointCustomShow {
        private readonly PowerPointDocument _presentation;
        private readonly PowerPointHeldNode _show;

        internal PowerPointCustomShow(PowerPointDocument presentation, PowerPointPartRef part, OfficeXmlNode show) {
            _presentation = presentation;
            _show = new PowerPointHeldNode(show, part);
        }

        internal OfficeXmlNode Node => _show.Element;

        /// <summary>Gets the stable custom-show identifier stored in the package.</summary>
        public uint Id => ReadId(Node) ?? 0U;

        internal static uint? ReadId(OfficeXmlNode show) =>
            show.Attribute("id") is string value ? XmlConvert.ToUInt32(value) : null;

        /// <summary>Gets the custom-show display name.</summary>
        public string Name => Node.Attribute("name") ?? string.Empty;

        /// <summary>Gets the slides in playback order.</summary>
        public IReadOnlyList<PowerPointSlide> Slides => _presentation.ResolveCustomShowSlides(this);

        /// <summary>Replaces the custom-show slide sequence.</summary>
        public PowerPointCustomShow SetSlides(IEnumerable<PowerPointSlide> slides) {
            _presentation.SetCustomShowSlides(this, slides);
            return this;
        }

        /// <summary>Adds a slide to the end of the custom show.</summary>
        public PowerPointCustomShow AddSlide(PowerPointSlide slide) => InsertSlide(Slides.Count, slide);

        /// <summary>Inserts a slide at the specified zero-based custom-show position.</summary>
        public PowerPointCustomShow InsertSlide(int index, PowerPointSlide slide) {
            _presentation.InsertCustomShowSlide(this, index, slide);
            return this;
        }

        /// <summary>Removes the first matching slide from the custom show.</summary>
        public bool RemoveSlide(PowerPointSlide slide) => _presentation.RemoveCustomShowSlide(this, slide);

        /// <summary>Moves a slide between zero-based custom-show positions.</summary>
        public PowerPointCustomShow MoveSlide(int sourceIndex, int destinationIndex) {
            _presentation.MoveCustomShowSlide(this, sourceIndex, destinationIndex);
            return this;
        }
    }

    public sealed partial class PowerPointDocument {
        private const int CustomShowNamespace = OfficeXmlNamespaces.PresentationML;

        /// <summary>Gets the custom shows in package order.</summary>
        public IReadOnlyList<PowerPointCustomShow> CustomShows {
            get {
                ThrowIfDisposed();
                PowerPointPartRef part = PresentationPartRef;
                var shows = new List<PowerPointCustomShow>();
                foreach (OfficeXmlNode show in CustomShowNodes(part)) {
                    shows.Add(new PowerPointCustomShow(this, part, show));
                }
                return new ReadOnlyCollection<PowerPointCustomShow>(shows);
            }
        }

        /// <summary>Creates a custom show with an explicit, nonempty slide sequence.</summary>
        public PowerPointCustomShow AddCustomShow(string name, IEnumerable<PowerPointSlide> slides) {
            ThrowIfDisposed();
            PowerPointPartRef part = CustomShowPartForEdit();
            string normalizedName = ValidateCustomShowName(part, name, default);
            PowerPointSlide[] resolvedSlides = ValidateCustomShowSlides(slides);
            part = part.Current;
            uint id = AllocateCustomShowId(part);
            string[] relationships = resolvedSlides.Select(slide => GetCustomShowSlideRelationship(part, slide)).ToArray();
            OfficeXmlNode root = part.Root;
            OfficeXmlNode list = KnownCustomShowChild(root, "custShowLst");
            if (list.IsNone) list = root.AddOrdered(CustomShowNamespace, "custShowLst", PPresentationView.Order);
            OfficeXmlNode show = list.CreateElement(CustomShowNamespace, "custShow");
            show.SetUInt32Attribute("id", id);
            show.SetAttribute("name", normalizedName);
            OfficeXmlNode slideList = show.CreateElement(CustomShowNamespace, "sldLst");
            foreach (string idOfSlide in relationships) slideList.Append(CreateCustomShowEntry(slideList, idOfSlide));
            show.Append(slideList);
            show = list.Append(show);
            if (!part.IsSdk) part.MarkEdited();
            return new PowerPointCustomShow(this, part, show);
        }

        /// <summary>Finds a custom show by name using ordinal, case-insensitive matching.</summary>
        public PowerPointCustomShow? GetCustomShow(string name) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(name)) return null;
            PowerPointPartRef part = PresentationPartRef;
            string lookupName = name.Trim();
            foreach (OfficeXmlNode show in CustomShowNodes(part)) {
                if (string.Equals(show.Attribute("name"), lookupName, StringComparison.OrdinalIgnoreCase)) {
                    return new PowerPointCustomShow(this, part, show);
                }
            }
            return null;
        }

        /// <summary>Renames a custom show while preserving its identifier and linked actions.</summary>
        public PowerPointCustomShow RenameCustomShow(PowerPointCustomShow customShow, string name) {
            ThrowIfDisposed();
            PowerPointPartRef part = CustomShowPartForEdit();
            OfficeXmlNode owned = RequireOwnedCustomShow(part, customShow);
            owned.SetAttribute("name", ValidateCustomShowName(part, name, owned));
            if (!part.IsSdk) part.MarkEdited();
            return customShow;
        }

        /// <summary>Removes a custom show and actions that target its identifier.</summary>
        public bool RemoveCustomShow(PowerPointCustomShow customShow) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(customShow);
            // Linked-action cleanup visits the package's typed parts. Bind the show before removing it.
            OpenSdkForPackageCompatibility();
            PowerPointPartRef part = PresentationPartRef;
            OfficeXmlNode owned = RequireOwnedCustomShow(part, customShow);
            uint? id = PowerPointCustomShow.ReadId(owned);
            OfficeXmlNode list = owned.Parent;
            owned.Remove();
            if (KnownCustomShowChild(list, "custShow").IsNone) list.Remove();
            if (id.HasValue) RemoveCustomShowLinks(id.Value);
            return true;
        }

        private PowerPointPartRef CustomShowPartForEdit() {
            if (!IsSdkLoaded && !CanEditPackage) OpenSdkForPackageCompatibility();
            return PresentationPartRef.ForMovingXmlNodes();
        }

        private static IEnumerable<OfficeXmlNode> CustomShowNodes(PowerPointPartRef part) =>
            KnownCustomShowChildren(KnownCustomShowChild(part.Root, "custShowLst"), "custShow");

        private static OfficeXmlNode KnownCustomShowChild(OfficeXmlNode parent, string name) {
            foreach (OfficeXmlNode child in parent.Children(CustomShowNamespace, name)) {
                if (!child.IsSdkUnknownElement) return child;
            }
            return default;
        }

        private static IEnumerable<OfficeXmlNode> KnownCustomShowChildren(OfficeXmlNode parent, string name) {
            foreach (OfficeXmlNode child in parent.Children(CustomShowNamespace, name)) {
                if (!child.IsSdkUnknownElement) yield return child;
            }
        }

        internal IReadOnlyList<PowerPointSlide> ResolveCustomShowSlides(PowerPointCustomShow show) {
            IReadOnlyList<PowerPointSlide> slides = CustomShowPresentationSlides();
            return ResolveCustomShowEntries(PresentationPartRef, show.Node, slides).Select(item => item.Slide).ToList().AsReadOnly();
        }

        private List<PowerPointSlide> CustomShowPresentationSlides() => GetPackageSlides() ?? LoadedSlides;

        private static List<(OfficeXmlNode Entry, PowerPointSlide Slide)> ResolveCustomShowEntries(PowerPointPartRef part,
            OfficeXmlNode show, IReadOnlyList<PowerPointSlide> slides) {
            var entries = new List<(OfficeXmlNode, PowerPointSlide)>();
            OfficeXmlNode slideList = KnownCustomShowChild(show, "sldLst");
            foreach (OfficeXmlNode entry in KnownCustomShowChildren(slideList, "sld")) {
                string? id = entry.Attribute(OfficeXmlNamespaces.Relationships, "id");
                if (string.IsNullOrEmpty(id) || !part.TryGetRelationship(id!, out PowerPointRelationship relationship) ||
                    relationship.IsExternal || relationship.TargetPart.Kind != PowerPointPartKind.Slide ||
                    OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type) != PowerPointPartRef.SlideRelationshipType) continue;
                foreach (PowerPointSlide slide in slides) {
                    if (!slide.PartRef.Equals(relationship.TargetPart)) continue;
                    entries.Add((entry, slide));
                    break;
                }
            }
            return entries;
        }

        internal void InsertCustomShowSlide(PowerPointCustomShow show, int index, PowerPointSlide slide) {
            ThrowIfDisposed();
            IReadOnlyList<PowerPointSlide> slides = CustomShowPresentationSlides();
            PowerPointPartRef part = CustomShowPartForEdit();
            OfficeXmlNode owned = RequireOwnedCustomShow(part, show);
            PowerPointSlide resolved = ValidateCustomShowSlides(new[] { slide })[0];
            List<(OfficeXmlNode Entry, PowerPointSlide Slide)> entries = ResolveCustomShowEntries(part, owned, slides);
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(index, entries.Count);
            OfficeXmlNode slideList = KnownCustomShowChild(owned, "sldLst");
            bool attach = slideList.IsNone;
            if (attach) slideList = owned.CreateElement(CustomShowNamespace, "sldLst");
            InsertCustomShowEntry(slideList, CreateCustomShowEntry(part, slideList, resolved), index, entries);
            if (attach) owned.Append(slideList);
            if (!part.IsSdk) part.MarkEdited();
        }

        internal bool RemoveCustomShowSlide(PowerPointCustomShow show, PowerPointSlide slide) {
            ThrowIfDisposed();
            IReadOnlyList<PowerPointSlide> slides = CustomShowPresentationSlides();
            PowerPointPartRef part = CustomShowPartForEdit();
            OfficeXmlNode owned = RequireOwnedCustomShow(part, show);
            ArgumentNullException.ThrowIfNull(slide);
            List<(OfficeXmlNode Entry, PowerPointSlide Slide)> entries = ResolveCustomShowEntries(part, owned, slides);
            OfficeXmlNode match = entries.FirstOrDefault(item => ReferenceEquals(item.Slide, slide)).Entry;
            if (match.IsNone) return false;
            if (entries.Count == 1) throw new InvalidOperationException("A custom show requires at least one slide.");
            match.Remove();
            if (!part.IsSdk) part.MarkEdited();
            return true;
        }

        internal void MoveCustomShowSlide(PowerPointCustomShow show, int sourceIndex, int destinationIndex) {
            ThrowIfDisposed();
            IReadOnlyList<PowerPointSlide> slides = CustomShowPresentationSlides();
            PowerPointPartRef part = CustomShowPartForEdit();
            OfficeXmlNode owned = RequireOwnedCustomShow(part, show);
            List<(OfficeXmlNode Entry, PowerPointSlide Slide)> entries = ResolveCustomShowEntries(part, owned, slides);
            ArgumentOutOfRangeException.ThrowIfNegative(sourceIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(sourceIndex, entries.Count);
            ArgumentOutOfRangeException.ThrowIfNegative(destinationIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(destinationIndex, entries.Count);
            if (sourceIndex == destinationIndex) return;
            OfficeXmlNode moved = entries[sourceIndex].Entry;
            moved.Remove();
            entries.RemoveAt(sourceIndex);
            InsertCustomShowEntry(KnownCustomShowChild(owned, "sldLst"), moved, destinationIndex, entries);
            if (!part.IsSdk) part.MarkEdited();
        }

        private static void InsertCustomShowEntry(OfficeXmlNode slideList, OfficeXmlNode entry, int index,
            List<(OfficeXmlNode Entry, PowerPointSlide Slide)> entries) {
            if (index < entries.Count) {
                slideList.InsertBefore(entry, entries[index].Entry);
            } else if (entries.Count > 0) {
                slideList.InsertBefore(entry, entries[entries.Count - 1].Entry.NextNode);
            } else {
                OfficeXmlNode reference = FirstCustomShowChild(slideList);
                while (reference.Is(CustomShowNamespace, "sld") && !reference.IsSdkUnknownElement) reference = reference.NextNode;
                slideList.InsertBefore(entry, reference);
            }
        }

        internal void SetCustomShowSlides(PowerPointCustomShow show, IEnumerable<PowerPointSlide> slides) {
            ThrowIfDisposed();
            PowerPointPartRef part = CustomShowPartForEdit();
            RequireOwnedCustomShow(part, show);
            PowerPointSlide[] resolved = ValidateCustomShowSlides(slides);
            // Enumerating caller-supplied slides may open the SDK and rebind the show.
            part = part.Current;
            OfficeXmlNode owned = show.Node;
            OfficeXmlNode slideList = KnownCustomShowChild(owned, "sldLst");
            bool attach = slideList.IsNone;
            if (attach) slideList = owned.CreateElement(CustomShowNamespace, "sldLst");
            List<OfficeXmlNode> existing = KnownCustomShowChildren(slideList, "sld").ToList();
            var entriesByRelationship = new Dictionary<string, Queue<OfficeXmlNode>>(StringComparer.Ordinal);
            foreach (OfficeXmlNode entry in existing) {
                string? id = entry.Attribute(OfficeXmlNamespaces.Relationships, "id");
                if (string.IsNullOrEmpty(id)) continue;
                if (!entriesByRelationship.TryGetValue(id!, out Queue<OfficeXmlNode>? entries)) {
                    entries = new Queue<OfficeXmlNode>();
                    entriesByRelationship.Add(id!, entries);
                }
                entries.Enqueue(entry);
            }

            var selected = new List<OfficeXmlNode>(resolved.Length);
            foreach (PowerPointSlide slide in resolved) {
                string id = GetCustomShowSlideRelationship(part, slide);
                selected.Add(entriesByRelationship.TryGetValue(id, out Queue<OfficeXmlNode>? entries) && entries.Count > 0
                    ? entries.Dequeue()
                    : CreateCustomShowEntry(slideList, id));
            }

            foreach (OfficeXmlNode entry in existing) entry.Remove();
            OfficeXmlNode first = FirstCustomShowChild(slideList);
            foreach (OfficeXmlNode entry in selected) slideList.InsertBefore(entry, first);
            if (attach) owned.Append(slideList);
            if (!part.IsSdk) part.MarkEdited();
        }

        private static OfficeXmlNode FirstCustomShowChild(OfficeXmlNode list) =>
            list.Tree != null ? list.FirstChild : list.ChildNodes().FirstOrDefault();

        private static OfficeXmlNode RequireOwnedCustomShow(PowerPointPartRef part, PowerPointCustomShow show) {
            ArgumentNullException.ThrowIfNull(show);
            OfficeXmlNode owned = show.Node;
            foreach (OfficeXmlNode current in CustomShowNodes(part)) {
                if (current.Equals(owned)) return owned;
            }
            throw new InvalidOperationException("The custom show does not belong to this presentation.");
        }

        private static string ValidateCustomShowName(PowerPointPartRef part, string name, OfficeXmlNode except) {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            string normalized = name.Trim();
            PowerPointXmlValueValidator.ValidateCharacters(normalized, nameof(name), "Custom-show name");
            foreach (OfficeXmlNode show in CustomShowNodes(part)) {
                if (!show.Equals(except) && string.Equals(show.Attribute("name"), normalized, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException($"A custom show named '{normalized}' already exists.");
                }
            }
            return normalized;
        }

        private PowerPointSlide[] ValidateCustomShowSlides(IEnumerable<PowerPointSlide> slides) {
            ArgumentNullException.ThrowIfNull(slides);
            PowerPointSlide[] resolved = slides.ToArray();
            if (resolved.Length == 0) throw new ArgumentException("A custom show requires at least one slide.", nameof(slides));
            List<PowerPointSlide> local = CustomShowPresentationSlides();
            if (resolved.Any(slide => slide == null || !local.Contains(slide)))
                throw new InvalidOperationException("Every custom-show slide must belong to this presentation.");
            return resolved;
        }

        private static OfficeXmlNode CreateCustomShowEntry(PowerPointPartRef part, OfficeXmlNode slideList, PowerPointSlide slide) =>
            CreateCustomShowEntry(slideList, GetCustomShowSlideRelationship(part, slide));

        private static OfficeXmlNode CreateCustomShowEntry(OfficeXmlNode slideList, string id) {
            OfficeXmlNode entry = slideList.CreateElement(CustomShowNamespace, "sld");
            entry.SetAttribute(OfficeXmlNamespaces.Relationships, "id", id);
            return entry;
        }

        private static string GetCustomShowSlideRelationship(PowerPointPartRef part, PowerPointSlide slide) =>
            part.RelationshipIdOf(slide.PartRef, PowerPointPartRef.SlideRelationshipType)
            ?? throw new InvalidOperationException("The slide has no relationship in this presentation.");

        private static uint AllocateCustomShowId(PowerPointPartRef part) {
            var used = new HashSet<uint>();
            foreach (OfficeXmlNode show in CustomShowNodes(part)) {
                if (PowerPointCustomShow.ReadId(show) is uint id) used.Add(id);
            }
            return FindAvailableUInt32Id(used, 1U, "custom-show");
        }

        private static uint FindAvailableUInt32Id(
            ISet<uint> usedIds, uint firstCandidate, string identifierKind,
            uint maximumInclusive = uint.MaxValue) {
            uint candidate = firstCandidate <= maximumInclusive ? firstCandidate : 0U;
            uint initialCandidate = candidate;
            do {
                if (!usedIds.Contains(candidate)) return candidate;
                candidate = candidate == maximumInclusive ? 0U : candidate + 1U;
            } while (candidate != initialCandidate);
            throw new InvalidOperationException($"The presentation has no available {identifierKind} identifiers.");
        }
    }
}
