using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPackageIntegrity {
        [Fact]
        public void GraphIsIntact() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using PresentationDocument document = PresentationDocument.Open(filePath, false);
            List<string> warnings = PptxDoctor.GetWarnings(document);
            Assert.Empty(warnings);
            File.Delete(filePath);
        }
    }

    internal static class PptxDoctor {
        internal static List<string> GetWarnings(PresentationDocument doc) {
            PresentationPart pPart = doc.PresentationPart!;
            Presentation pres = pPart.Presentation!;
            List<string> warnings = new();

            if (pres.SlideMasterIdList == null || !pres.SlideMasterIdList.Elements<SlideMasterId>().Any()) {
                warnings.Add("Missing or empty <p:sldMasterIdLst> (no master attached)");
            }

            if (pres.SlideIdList == null || !pres.SlideIdList.Elements<SlideId>().Any()) {
                warnings.Add("Missing or empty <p:sldIdLst> (no slides registered)");
            }

            HashSet<uint> seenIds = new();
            foreach (SlideId sldId in pres.SlideIdList?.Elements<SlideId>() ?? Enumerable.Empty<SlideId>()) {
                if (!seenIds.Add(sldId.Id!.Value)) {
                    warnings.Add($"Duplicate <p:sldId/@Id> {sldId.Id.Value} (must be unique).");
                }

                if (pPart.GetReferenceRelationship(sldId.RelationshipId!) == null) {
                    warnings.Add($"Presentation rel '{sldId.RelationshipId}' not found for a slide.");
                }
            }

            foreach (SlideMasterPart master in pPart.SlideMasterParts) {
                if (!master.SlideLayoutParts.Any()) {
                    warnings.Add("SlideMasterPart has no SlideLayoutPart (need at least one).");
                }

                if (master.ThemePart == null) {
                    warnings.Add("SlideMasterPart has no ThemePart (PowerPoint will generate/repair one).");
                }

                HashSet<string> layoutIds = master.SlideMaster!.SlideLayoutIdList!
                    .Elements<SlideLayoutId>().Select(x => x.RelationshipId!).ToHashSet();
                HashSet<string> actualLayoutRelIds = master
                    .GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
                    .Select(r => r.Id).ToHashSet();

                if (!layoutIds.SetEquals(actualLayoutRelIds)) {
                    warnings.Add("slideMaster layout id list doesn’t match its rels (broken master→layout links).");
                }
            }

            foreach (SlidePart slide in pPart.SlideParts) {
                bool hasLayoutRel = slide
                    .GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")
                    .Any();
                if (!hasLayoutRel) {
                    warnings.Add("SlidePart without a slideLayout relationship.");
                }

                if (slide.Slide?.CommonSlideData?.ShapeTree == null) {
                    warnings.Add("Slide is missing <p:cSld>/<p:spTree> (minimal content).");
                }
            }

            if (pres.SlideSize == null) {
                warnings.Add("Missing <p:sldSz> (slide size).");
            }

            if (pres.NotesSize == null) {
                warnings.Add("Missing <p:notesSz> (notes size).");
            }

            return warnings;
        }
    }
}
