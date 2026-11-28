using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ProjectLegacyCustomShows(LegacyPptImportParts parts,
            LegacyPptDocument legacy,
            IReadOnlyDictionary<uint, PowerPointPartRef> slidePartsByLegacyId) {
            if (legacy.CustomShows.Count == 0) return;
            PowerPointPartRef presentationPart = parts.PresentationPart;
            OfficeXmlNode root = presentationPart.Root;
            OfficeXmlNode list = root.CreateElement(LegacyPptXml.P, "custShowLst");
            uint customShowId = 0;
            bool hasChildren = false;
            foreach (LegacyPptCustomShow source in legacy.CustomShows) {
                if (source.Name.Length == 0) continue;
                OfficeXmlNode show = list.AddP("custShow")
                    .Set("name", source.Name)
                    .Set("id", (long)++customShowId);
                OfficeXmlNode slideList = show.AddP("sldLst");
                foreach (uint slideId in source.SlideIds) {
                    if (!slidePartsByLegacyId.TryGetValue(slideId,
                            out PowerPointPartRef slidePart)) continue;
                    slideList.AddP("sld").SetRelationship("id", parts.RelationshipIdTo(presentationPart, slidePart));
                }
                hasChildren = true;
            }
            if (!hasChildren) return;
            OfficeXmlNode existing = root.Child(LegacyPptXml.P, "custShowLst");
            OfficeXmlNode placed;
            if (existing.IsNone) {
                placed = root.InsertOrdered(list, PPresentationView.Order);
            } else {
                placed = root.InsertBefore(list, existing);
                existing.Remove();
            }
            presentationPart.MarkEdited();
            parts.CustomShowIds = CreateLegacyCustomShowIdIndex(placed);
        }

        // Custom show ids by name; a name several shows share has no id, as a hyperlink to it can't tell them apart.
        private static Dictionary<string, uint?> CreateLegacyCustomShowIdIndex(OfficeXmlNode list) =>
            list.Children(LegacyPptXml.P, "custShow").ToList()
                .GroupBy(show => show.Attribute("name") ?? string.Empty,
                    StringComparer.Ordinal)
                .ToDictionary(group => group.Key,
                    group => group.Count() == 1
                        ? group.First().UInt("id")
                        : null,
                    StringComparer.Ordinal);
    }
}
