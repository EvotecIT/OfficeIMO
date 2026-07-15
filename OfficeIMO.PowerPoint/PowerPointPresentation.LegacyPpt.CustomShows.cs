using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyCustomShows(PowerPointPresentation presentation,
            LegacyPptPresentation legacy,
            IReadOnlyDictionary<uint, SlidePart> slidePartsByLegacyId) {
            if (legacy.CustomShows.Count == 0) return;
            PresentationPart presentationPart = presentation._presentationPart;
            var list = new P.CustomShowList();
            uint customShowId = 0;
            foreach (LegacyPptCustomShow source in legacy.CustomShows) {
                if (source.Name.Length == 0) continue;
                var slideList = new P.SlideList();
                foreach (uint slideId in source.SlideIds) {
                    if (!slidePartsByLegacyId.TryGetValue(slideId,
                            out SlidePart? slidePart)) continue;
                    slideList.Append(new P.SlideListEntry {
                        Id = presentationPart.GetIdOfPart(slidePart)
                    });
                }
                list.Append(new P.CustomShow(slideList) {
                    Name = source.Name,
                    Id = ++customShowId
                });
            }
            if (list.HasChildren) {
                presentationPart.Presentation!.CustomShowList = list;
            }
        }
    }
}
