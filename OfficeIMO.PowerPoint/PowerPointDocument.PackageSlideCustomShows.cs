using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private sealed record SlideCustomShowPruning(
            OfficeXmlNode List, List<OfficeXmlNode> Entries, List<OfficeXmlNode> EmptyShows, List<CustomShowLinkEdit>? LinkEdits);

        private bool TryPlanPackageSlideCustomShows(OfficeOpcEditablePackage package, PackageSlideList slides, string relationship,
            out SlideCustomShowPruning? plan) {
            plan = null;
            OfficeXmlTree presentation = slides.Presentation;
            int listNode = presentation.FindChild(presentation.Root, CustomShowNamespace, "custShowLst");
            if (listNode == OfficeXmlTree.NoNode) return true;
            if (presentation.HasCommentsOrInstructions()) return false;
            var list = new OfficeXmlNode(presentation, listNode);

            // With alternate markup the SDK may recognize entries beneath the wrapper. Use its existing
            // cleanup if the affected relationship reaches a candidate there.
            foreach (OfficeXmlNode alternate in list.Descendants(OfficeXmlNamespaces.MarkupCompatibility, "AlternateContent")) {
                if (IsInsideCustomShowExtension(alternate, list)) continue;
                foreach (OfficeXmlNode entry in alternate.ContentDescendants(CustomShowNamespace, "sld")) {
                    if (CustomShowEntryMatchesSlide(entry, relationship)) return false;
                }
            }

            var entries = new List<OfficeXmlNode>();
            var empty = new List<OfficeXmlNode>();
            var emptiedIds = new HashSet<uint>();
            foreach (OfficeXmlNode show in KnownCustomShowChildren(list, "custShow")) {
                OfficeXmlNode slideList = KnownCustomShowChild(show, "sldLst");
                // The SDK visits all typed descendants, including additional lists from malformed producers.
                if (slideList.NextSiblingNamed(CustomShowNamespace, "sldLst").Exists) return false;
                int before = entries.Count;
                bool retainsSlide = false;
                foreach (OfficeXmlNode entry in KnownCustomShowChildren(slideList, "sld")) {
                    if (CustomShowEntryMatchesSlide(entry, relationship)) entries.Add(entry);
                    else retainsSlide = true;
                }
                if (before == entries.Count || retainsSlide) continue;

                empty.Add(show);
                try {
                    if (PowerPointCustomShow.ReadId(show) is uint id) emptiedIds.Add(id);
                } catch (Exception exception) when (exception is FormatException or OverflowException) {
                    return false;
                }
            }

            List<CustomShowLinkEdit>? links = null;
            if (emptiedIds.Count != 0 &&
                !TryPlanOwnedCustomShowLinks(package, PowerPointPartRef.Package(this, slides.Main), emptiedIds, out links)) return false;
            plan = new SlideCustomShowPruning(list, entries, empty, links);
            return true;
        }

        private static bool IsInsideCustomShowExtension(OfficeXmlNode node, OfficeXmlNode list) {
            for (OfficeXmlNode parent = node.Parent; parent.Exists && parent != list; parent = parent.Parent) {
                if (parent.LocalName == "extLst") return true;
            }
            return false;
        }

        private static bool CustomShowEntryMatchesSlide(OfficeXmlNode entry, string relationship) =>
            string.Equals(entry.Attribute(OfficeXmlNamespaces.Relationships, "id"), relationship, StringComparison.Ordinal);

        private static void ApplyPackageSlideCustomShows(SlideCustomShowPruning plan) {
            if (plan.EmptyShows.Count != 0) plan.List.Tree!.RetainSourceBytes();
            foreach (OfficeXmlNode entry in plan.Entries) entry.Remove();
            foreach (OfficeXmlNode show in plan.EmptyShows) show.Remove();
            if (KnownCustomShowChild(plan.List, "custShow").IsNone) plan.List.Remove();
            ApplyOwnedCustomShowLinkEdits(plan.LinkEdits);
        }
    }
}
