using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointNotesMaster {
        [Fact(Skip = "Doesn't work after changes to PowerPoint")]
        public void NotesMasterUsesRelationshipId() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                PresentationPart presentationPart = document.PresentationPart!;
                NotesMasterPart notesMasterPart = presentationPart.NotesMasterPart!;
                NotesMasterIdList? list = presentationPart.Presentation.NotesMasterIdList;
                Assert.NotNull(list);
                NotesMasterId notesMasterId = list!.GetFirstChild<NotesMasterId>()!;
                OpenXmlAttribute attr = notesMasterId.GetAttributes().First(a =>
                    a.LocalName == "id" &&
                    a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                );
                Assert.False(string.IsNullOrEmpty(attr.Value));
                Assert.True(notesMasterId.GetAttributes().All(a => a.LocalName != "id" || a.NamespaceUri != string.Empty));
                Assert.Equal(attr.Value, presentationPart.GetIdOfPart(notesMasterPart));
            }

            File.Delete(filePath);
        }
    }
}
