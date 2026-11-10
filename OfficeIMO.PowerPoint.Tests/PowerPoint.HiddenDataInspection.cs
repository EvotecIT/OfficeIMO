using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointHiddenDataTests {
        [Fact]
        public void RemoveHiddenDataAnonymizesKeptCommentsAndRemovesNotesHiddenSlidesAndOffSlideShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide visible = presentation.AddSlide();
                    visible.AddTextBox("Agenda", OfficeLength.Points(40), OfficeLength.Points(40), OfficeLength.Points(300), OfficeLength.Points(40));
                    visible.AddTextBox("Parked idea", OfficeLength.Points(-400), OfficeLength.Points(40), OfficeLength.Points(200), OfficeLength.Points(40));
                    visible.Notes.Text = "Mention the confidential discount.";
                    var reviewer = new PowerPointCommentAuthor("Private Reviewer", "PR");
                    presentation.AddClassicComment(visible, reviewer, "Classic note.");
                    presentation.AddModernComment(visible, reviewer, "Modern note.");
                    PowerPointSlide hidden = presentation.AddSlide();
                    hidden.AddTextBox("Backup pricing", OfficeLength.Points(40), OfficeLength.Points(40), OfficeLength.Points(300), OfficeLength.Points(40));
                    hidden.Hide();
                    presentation.Properties.Creator = "Private Author";
                    presentation.Save();
                }

                int slidesBefore;
                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    slidesBefore = presentation.Slides.Count;
                    OfficeHiddenDataReport found = presentation.InspectHiddenData();
                    Assert.Equal(2, found.GetFindings(OfficeHiddenDataKind.Comments).Count);
                    Assert.Contains(found.GetFindings(OfficeHiddenDataKind.SpeakerNotes), finding => finding.Value == "Mention the confidential discount.");
                    Assert.Contains(found.GetFindings(OfficeHiddenDataKind.HiddenContent), finding => finding.Description == "Hidden slide");
                    Assert.Contains(found.GetFindings(OfficeHiddenDataKind.HiddenContent), finding => finding.Description == "Shape placed off the slide");

                    OfficeHiddenDataReport removed = presentation.RemoveHiddenData(new OfficeHiddenDataOptions {
                        Kinds = OfficeHiddenDataKind.All & ~OfficeHiddenDataKind.Comments,
                        AnonymousAuthor = "Reviewer"
                    });
                    Assert.All(removed.GetFindings(OfficeHiddenDataKind.Comments), finding => Assert.Equal(OfficeHiddenDataOutcome.Retained, finding.Outcome));
                    Assert.Contains(removed.GetFindings(OfficeHiddenDataKind.PersonalIdentities), finding => finding.Outcome == OfficeHiddenDataOutcome.Anonymized && finding.Value == "Private Reviewer");
                    presentation.Save();
                }

                using (PowerPointDocument reopened = PowerPointDocument.Load(filePath)) {
                    Assert.Empty(reopened.ValidateDocument());
                    Assert.Equal(slidesBefore - 1, reopened.Slides.Count);
                    PowerPointSlide slide = reopened.Slides.Single(candidate => candidate.Shapes.OfType<PowerPointTextBox>().Any(box => box.Text == "Agenda"));
                    Assert.DoesNotContain(slide.Shapes.OfType<PowerPointTextBox>(), box => box.Text == "Parked idea");
                    Assert.False(slide.HasSpeakerNotes);
                    Assert.Equal("Reviewer", Assert.Single(reopened.GetClassicComments(slide)).Author.Name);
                    Assert.Equal("Reviewer", Assert.Single(reopened.GetModernComments(slide)).Author.Name);

                    OfficeHiddenDataReport after = reopened.InspectHiddenData();
                    Assert.True(after.Kinds == (OfficeHiddenDataKind.Comments | OfficeHiddenDataKind.PersonalIdentities), string.Join(Environment.NewLine, after.Findings));
                    Assert.Equal("Reviewer", Assert.Single(after.GetFindings(OfficeHiddenDataKind.PersonalIdentities)).Value);
                }

                using PresentationDocument package = PresentationDocument.Open(filePath, false);
                string authors = string.Concat(package.PresentationPart!.Parts.Select(pair => pair.OpenXmlPart)
                    .Where(part => part.Uri.ToString().IndexOf("author", StringComparison.OrdinalIgnoreCase) >= 0)
                    .Select(part => part.RootElement?.OuterXml));
                Assert.DoesNotContain("Private Reviewer", authors);
            } finally {
                File.Delete(filePath);
            }
        }
    }
}
