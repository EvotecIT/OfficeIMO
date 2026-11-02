using OfficeIMO;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Exercises public native source saving on the generated three-page newsletter without native application calls.</summary>
internal static class SourceNewsletterLifecycle {
    internal static int Run(string inputPath, string replacementImagePath, string outputDirectory) {
        byte[] source = File.ReadAllBytes(inputPath), replacement = File.ReadAllBytes(replacementImagePath);
        string output = Path.GetFullPath(outputDirectory);
        if (Directory.Exists(output) && Directory.EnumerateFileSystemEntries(output).Any())
            throw new ArgumentException("Choose an empty lifecycle output directory.", nameof(outputDirectory));
        Directory.CreateDirectory(output);
        void Require(bool condition, string message) { if (!condition) throw new InvalidDataException(message); }
        string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        OfficeRichTextParagraph Rewrite(OfficeRichTextParagraph p, Func<string, string> change) {
            var runs = p.Runs.Select(r => new OfficeRichTextRun(change(r.Text), r.FontSize, r.Color, r.Bold, r.Italic, r.Underline,
                r.FontFamily, r.Strikethrough, r.BackgroundColor, r.UnderlineStyle, r.StrikethroughStyle, r.Baseline, r.ParagraphIndent)
                .WithLanguageId(r.LanguageId).WithAutomaticPairKerningThreshold(r.AutomaticPairKerningThreshold)).ToArray();
            var result = new OfficeRichTextParagraph(runs, p.Alignment, p.LineHeight, p.Margins, p.Indent, p.LineHeightFactor).WithTabStops(p.TabStops);
            if (p.SuppressAutomaticHyphenation.HasValue) result = result.WithAutomaticHyphenationSuppressed(p.SuppressAutomaticHyphenation.Value);
            return result;
        }
        object Topology(PublisherDocument d) => new {
            pages = d.Pages.Concat(d.MasterPages).Select(p => new { p.Id, p.MasterPageId, p.ObjectIds,
                frames = p.TextFrames.Select(f => new { f.Id, f.PageId, f.StoryId, f.X, f.Y, f.Width, f.Height, f.Order, f.NextFrameId, f.PreviousFrameId }),
                tables = p.Tables.Select(t => new { t.Id, t.PageId, t.StoryId, t.X, t.Y, t.Width, t.Height, t.ColumnWidths, t.RowHeights,
                    cells = t.Cells.Select(c => new { c.RowIndex, c.ColumnIndex, c.RowSpan, c.ColumnSpan }) }),
                pictures = p.Pictures.Select(i => new { i.Id, i.PageId, i.X, i.Y, i.Width, i.Height, i.Crop }) })
        };
        var cases = new List<object>();
        foreach (string kind in new[] { "linked", "table", "image", "combined" }) {
            PublisherDocument doc = PublisherDocument.Load(source);
            Require(doc.ToPubBytes().SequenceEqual(source), "Public exact-copy save changed the generated source.");
            string topology = JsonSerializer.Serialize(Topology(doc));
            PublisherTextStory body = doc.TextStories.Single(s => s.Text.Contains("Our community shares practical ideas"));
            PublisherTable table = doc.Pages[0].Tables.Single();
            PublisherPicture picture = doc.Pages[1].Pictures.Single(), logo = doc.MasterPages.Single().Pictures.Single();
            byte[] logoBytes = doc.Images.Single(i => i.Id == logo.ImageId).GetBytes();
            Require(picture.ImageId == logo.ImageId, "Input does not exercise one shared native asset.");
            bool linked = kind is "linked" or "combined", cells = kind is "table" or "combined", image = kind is "image" or "combined";
            if (linked) doc.ReplaceStoryParagraphs(body.Id, body.Paragraphs.Select(p => Rewrite(p, text => text.Replace("practical ideas", "helpful advice"))).ToArray());
            if (cells) {
                OfficeRichTextParagraph prior = table.Cells.Single(c => c.RowIndex == 1 && c.ColumnIndex == 1).Paragraphs.Single();
                doc.ReplaceTableCellParagraphs(table.Id, 1, 1, new[] { Rewrite(prior, _ => "Sunday"), Rewrite(prior, _ => "10:00") });
            }
            if (image) doc.ReplacePictureImage(picture.Id, replacement);
            var requestedStories = doc.TextStories.ToDictionary(s => s.Id, s => s.Text);
            string path = Path.Combine(output, kind + ".pub");
            bool strictRefused = false;
            try { doc.Save(path); } catch (OfficeConversionException) { strictRefused = true; }
            Require(strictRefused && !File.Exists(path), "Default strict save did not preserve the destination boundary.");
            PublisherSaveReport report = doc.Save(path, new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
            Require(!report.HasErrors && report.HasLoss && !report.IsExactSourceCopy, "Public save report lost the current unassessed-edit contract.");
            byte[] saved = File.ReadAllBytes(path);
            PublisherDocument actual = PublisherDocument.Load(saved);
            Require(JsonSerializer.Serialize(Topology(actual)) == topology, "Save changed unaffected frame/table/picture topology.");
            Require(actual.TextStories.Count == requestedStories.Count && actual.TextStories.All(s => requestedStories[s.Id] == s.Text), "Complete edited or neighboring stories changed.");
            Require(actual.ToPubBytes().SequenceEqual(saved), "Saved source cannot be copied exactly through the default public API.");
            Require(actual.Images.Single(i => i.Id == actual.MasterPages.Single().Pictures.Single().ImageId).GetBytes().SequenceEqual(logoBytes), "Replacing a placement also changed the shared master asset.");
            byte[] pageImage = actual.Images.Single(i => i.Id == actual.Pages[1].Pictures.Single().ImageId).GetBytes();
            Require(pageImage.SequenceEqual(image ? replacement : logoBytes), "Saved picture placement has the wrong embedded bytes.");
            Require(actual.Pages[0].Tables.Single().Cells.Select(c => c.Text).SequenceEqual(cells
                ? new[] { "Activity", "Date", "Community meeting", "Sunday\n10:00" }
                : new[] { "Activity", "Date", "Community meeting", "Saturday" }), "Cell CR or text partitions changed.");
            var item = new { kind, sha256 = Hash(saved), bytes = saved.Length, publicSave = true, strictDefaultRefused = strictRefused,
                expectedBodyText = actual.TextStories.Single(s => s.Id == body.Id).Text, expectedCellText = cells ? "Sunday\n10:00" : "Saturday",
                replacementJpeg = image, printStateNativeAcceptance = "not-run", fidelity = report.FidelityDiagnostics.Select(d => new { d.Code, d.LossKind, d.Source, d.Location }) };
            cases.Add(item); Console.WriteLine("PASS public source-save " + kind);
        }
        using var proof = new FileStream(Path.Combine(output, "public-save-proof.json"), FileMode.CreateNew, FileAccess.Write);
        JsonSerializer.Serialize(proof, new { sourceSha256 = Hash(source), replacementSha256 = Hash(replacement), cases }, new JsonSerializerOptions { WriteIndented = true });
        return 0;

    }
}
