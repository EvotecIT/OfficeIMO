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
        [Fact(Skip = "Doesn't work after changes to PowerPoint")]
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
                    if (master.SlideMaster == null) {
                        warnings.Add("Slide master missing root element.");
                    }

                    if (!master.SlideLayoutParts.Any()) {
                        warnings.Add("Slide master missing slide layout part.");
                    }

                    foreach (SlideLayoutPart layout in master.SlideLayoutParts) {
                        if (layout.SlideLayout == null) {
                            warnings.Add("Slide layout part missing root element.");
                        }
                    }

                    if (master.ThemePart == null) {
                        warnings.Add("Slide master missing theme part.");
                    } else if (master.ThemePart.Theme == null) {
                        warnings.Add("Theme part missing root element.");
                    }
                }

                if (presentationPart.ThemePart == null) {
                    warnings.Add("Presentation missing theme part.");
                } else if (presentationPart.ThemePart.Theme == null) {
                    warnings.Add("Presentation theme part missing root element.");
                }

                if (presentationPart.NotesMasterPart == null) {
                    warnings.Add("Presentation missing notes master part.");
                } else if (presentationPart.NotesMasterPart.NotesMaster == null) {
                    warnings.Add("Notes master part missing root element.");
                }

                ExtendedFilePropertiesPart? appPart = document.ExtendedFilePropertiesPart;
                if (appPart?.Properties == null) {
                    warnings.Add("Presentation app part missing root element.");
                }

                PresentationPropertiesPart? presPropsPart = presentationPart.PresentationPropertiesPart;
                if (presPropsPart?.PresentationProperties == null) {
                    warnings.Add("Presentation properties part missing root element.");
                }

                ViewPropertiesPart? viewPropsPart = presentationPart.ViewPropertiesPart;
                if (viewPropsPart?.ViewProperties == null) {
                    warnings.Add("View properties part missing root element.");
                }

                TableStylesPart? tableStylesPart = presentationPart.TableStylesPart;
                if (tableStylesPart?.TableStyleList == null) {
                    warnings.Add("Table styles part missing root element.");
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
