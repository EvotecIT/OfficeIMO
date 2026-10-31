using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors the independent multiple-master controls through the canonical public page and story model.</summary>
internal static class CandidateMasterPages {
    internal static CandidatePublicationSpec Create(bool paired) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage centered = document.AddMasterPage("Master Page A");
        PublisherPage[] masters;
        int[] assignments;
        if (paired) {
            var pair = document.AddMasterPagePair("Comparison Two-Page Master");
            masters = new[] { centered, pair.Left, pair.Right };
            assignments = new[] { 0, 1, 2 };
        } else {
            masters = new[] { centered, document.AddMasterPage("Comparison Master B") };
            assignments = new[] { 0, 1, 0 };
        }
        document.SetPageMaster(document.Pages[0].Id, masters[assignments[0]].Id);
        document.AddPage(masters[assignments[1]].Id);
        document.AddPage(masters[assignments[2]].Id);
        var objects = new List<CandidateObjectSpec>();
        for (int index = 0; index < masters.Length; index++) {
            string text = paired ? new[] { "Default master", "Paired master left", "Paired master right" }[index]
                : index == 0 ? "Master Alpha" : "Master Beta";
            Text(masters[index].Id, index, true, text, paired ? index == 2 ? 342 : 18 : index == 1 ? 342 : 18, 24, 252, 24, 10);
        }
        for (int index = 0; index < document.Pages.Count; index++)
            Text(document.Pages[index].Id, index, false, (paired ? "TwoPageMaster" : "MultipleMasters") + " page " + (index + 1), 72, 96, 468, 48, 16);
        CandidateMasterSpec[] definitions = masters.Select((master, index) => new CandidateMasterSpec(
            index == 0 ? "A" : paired ? "P" : "B", master.Name, master.MasterSide!.Value)).ToArray();
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count, definitions,
            document.Pages.Select(page => (int?)Array.FindIndex(masters, master => master.Id == page.MasterPageId)).ToArray());

        void Text(uint owner, int ownerIndex, bool onMaster, string text, double x, double y, double width, double height, double size) {
            PublisherTextFrame frame = document.AddTextFrame(owner, x, y, width, height, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") })
            });
            objects.Add(new CandidateObjectSpec(CandidateObjectKind.TextBox, new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == frame.StoryId).Paragraphs.Single()),
                PageIndex: ownerIndex, OnMaster: onMaster));
        }
    }
}
