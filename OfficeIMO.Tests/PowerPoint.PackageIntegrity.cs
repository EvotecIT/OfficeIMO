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
        public void CanValidatePackageStructure() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            List<string> warnings = PptxDoctor(filePath);

            Assert.True(warnings.Count == 0, string.Join(Environment.NewLine, warnings));

            File.Delete(filePath);
        }

        private static List<string> PptxDoctor(string path) {
            List<string> warnings = new();

            using (PresentationDocument document = PresentationDocument.Open(path, false)) {
                PresentationPart presentationPart = document.PresentationPart!;

                if (!presentationPart.SlideMasterParts.Any()) {
                    warnings.Add("Presentation missing slide master part.");
                }

                foreach (SlideMasterPart master in presentationPart.SlideMasterParts) {
                    if (!master.SlideLayoutParts.Any()) {
                        warnings.Add("Slide master missing slide layout part.");
                    }

                    if (master.ThemePart == null) {
                        warnings.Add("Slide master missing theme part.");
                    }
                }

                if (presentationPart.ThemePart == null) {
                    warnings.Add("Presentation missing theme part.");
                }

                List<SlidePart> slideParts = presentationPart.SlideParts.ToList();
                if (slideParts.Count == 0) {
                    warnings.Add("Presentation missing slide parts.");
                }

                SlideIdList? slideIdList = presentationPart.Presentation.SlideIdList;
                if (slideIdList != null) {
                    List<SlideId> slideIds = slideIdList.Elements<SlideId>().ToList();
                    List<uint> ids = slideIds.Select(s => s.Id!.Value).ToList();
                    if (ids.Count != ids.Distinct().Count()) {
                        warnings.Add("Duplicate slide identifiers detected.");
                    }

                    foreach (SlideId slideId in slideIds) {
                        string relId = slideId.RelationshipId!;
                        bool exists = presentationPart.Parts.Any(p => p.RelationshipId == relId && p.OpenXmlPart is SlidePart);
                        if (!exists) {
                            warnings.Add($"Slide id {slideId.Id!.Value} has invalid relationship id {relId}.");
                        }
                    }

                    List<string> slideRelIds = slideIds.Select(s => s.RelationshipId!.Value!).ToList();
                    foreach (SlidePart slidePart in slideParts) {
                        string? relId = presentationPart.GetIdOfPart(slidePart);
                        if (relId == null || !slideRelIds.Contains(relId)) {
                            warnings.Add("Slide part not referenced by presentation.");
                        }

                        if (slidePart.SlideLayoutPart == null) {
                            warnings.Add($"Slide {relId ?? "<unknown>"} missing layout relationship.");
                        }
                    }
                } else {
                    warnings.Add("Presentation missing slide id list.");
                }
            }

            return warnings;
        }
    }
}
