using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Publisher.Internal;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherMasterAssignmentTests {
    private static PublisherSaveOptions Allow => new() { LossPolicy = OfficeConversionLossPolicy.Allow };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Clearing_a_zero_or_absent_reference_is_a_byte_exact_noop(bool absent) {
        byte[] source = PublisherInputContractTests.Mutate("Contents", bytes => {
            var context = new PublisherParseContext(new PublisherReadOptions(), default);
            var data = new PublisherBinaryData(bytes, "Contents");
            PublisherContents model = new PublisherContentsReader(data, context).Read();
            PublisherContentChunk chunk = model.Chunks[model.Pages.First(page => !page.IsMaster).Id];
            PublisherBlock reference = new PublisherBlockReader(data, context).Chunk(chunk.Offset, chunk.End).Single(field => field.Id == 0x0D);
            if (absent) bytes[reference.Offset - 2] = 0x1F;
            else Buffer.BlockCopy(BitConverter.GetBytes(0U), 0, bytes, reference.Offset, 4);
        }, "NativeComparison/MultipleMasters/native-multiplemasters.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Null(document.Pages[0].MasterPageId);
        document.SetPageMaster(document.Pages[0].Id, null);
        var result = document.ToPubBytesResult(Allow);
        Assert.True(result.Report.IsExactSourceCopy); Assert.Equal(source, result.Value);
    }

    [Theory]
    [InlineData("MultipleMastersIgnored", "MultipleMasters", "Master Alpha")]
    [InlineData("TwoPageMasterIgnored", "TwoPageMaster", "Default master")]
    public void Native_ignored_master_is_not_applied_and_noop_native_copy_retains_latent_reference(string fixture, string body, string hidden) {
        byte[] source = File.ReadAllBytes(PathFor(fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Null(document.Pages[0].MasterPageId); Assert.Null(document.Pages[2].MasterPageId);
        Assert.Equal(document.MasterPages[1].Id, document.Pages[1].MasterPageId);
        Assert.All(document.Pages, page => Assert.Single(page.TextFrames));
        Assert.All(document.MasterPages, page => Assert.Single(page.TextFrames));
        Assert.DoesNotContain(hidden, document.ToSvg(0));
        Assert.Contains(body + " page 1", document.ToSvg(0));
        PdfReadDocument pdf = PdfReadDocument.Open(document.ToPdfBytes());
        Assert.DoesNotContain(hidden, pdf.Pages[0].ExtractText());
        Assert.Contains(body + " page 1", pdf.Pages[0].ExtractText());
        Assert.Contains(document.TextStories, story => story.Text.StartsWith(hidden, StringComparison.Ordinal));
        foreach (PublisherPage page in document.Pages) document.SetPageMaster(page.Id, page.MasterPageId);
        var saved = document.ToPubBytesResult(Allow);
        Assert.True(saved.Report.IsExactSourceCopy); Assert.Equal(source, saved.Value);
        var parsed = Contents(source);
        Assert.NotNull(parsed.Model.Pages.First(page => !page.IsMaster).Master);
        Assert.False(parsed.Model.Pages.First(page => !page.IsMaster).DisplayMaster);
    }

    [Theory]
    [InlineData("MultipleMasters/native-multiplemasters.pub", false)]
    [InlineData("TwoPageMaster/native-twopagemaster.pub", true)]
    [InlineData("GeneratedMultipleMasters/generated-multiplemasters.pub", false)]
    [InlineData("GeneratedMultipleMasters/native-normalized.pub", false)]
    [InlineData("GeneratedTwoPageMaster/generated-twopagemaster.pub", true)]
    [InlineData("GeneratedTwoPageMaster/native-normalized.pub", true)]
    [InlineData("MultipleMastersIgnored/native-multiplemastersignored.pub", false)]
    [InlineData("TwoPageMasterIgnored/native-twopagemasterignored.pub", true)]
    public void Assign_clear_and_restore_existing_masters_preserve_all_other_records_through_repeated_native_saves(string fixture, bool paired) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/" + fixture));
        PublisherDocument document = PublisherDocument.Load(source);
        uint?[] original = document.Pages.Select(page => page.MasterPageId).ToArray();
        uint[] masters = document.MasterPages.Select(page => page.Id).ToArray();
        var sides = document.MasterPages.Select(page => page.MasterSide).ToArray();
        var stories = document.TextStories.Select(story => (story.Id, story.Text)).ToArray();
        uint?[][] stages = { new uint?[] { masters[paired ? 2 : 1], masters[paired ? 1 : 0], masters[paired ? 0 : 1] },
            new uint?[] { null, masters[1], null }, original };
        foreach (uint?[] requested in stages) {
            var detached = document.Pages;
            uint?[] before = detached.Select(page => page.MasterPageId).ToArray();
            for (int index = 0; index < requested.Length; index++) document.SetPageMaster(document.Pages[index].Id, requested[index]);
            Assert.Equal(before, detached.Select(page => page.MasterPageId));
            var saved = document.ToPubBytesResult(Allow);
            Assert.False(saved.Report.HasErrors);
            if (!before.SequenceEqual(requested)) {
                Assert.False(saved.Report.IsExactSourceCopy);
                Assert.Contains(saved.Report.FidelityDiagnostics, value => value.Code == "PUB_NATIVE_MASTER_ASSIGNMENT_UNQUALIFIED");
            }
            PublisherNativeWritingTests.AssertPreservedStorage(source, saved.Value, "Contents");
            AssertUnrelatedContents(source, saved.Value);
            PublisherDocument reopened = PublisherDocument.Load(saved.Value);
            Assert.Equal(requested, reopened.Pages.Select(page => page.MasterPageId));
            Assert.Equal(masters, reopened.MasterPages.Select(page => page.Id));
            Assert.Equal(sides, reopened.MasterPages.Select(page => page.MasterSide));
            Assert.Equal(stories, reopened.TextStories.Select(story => (story.Id, story.Text)));
            Assert.All(reopened.Pages, page => Assert.Single(page.TextFrames));
            foreach (PublisherPage page in reopened.Pages) {
                string paint = PublisherAuthoringTests.PaintedText(page);
                foreach (PublisherPage master in reopened.MasterPages) {
                    string marker = reopened.TextStories.Single(story => story.Id == master.TextFrames[0].StoryId).Text.TrimEnd('\n');
                    if (page.MasterPageId == master.Id) Assert.Contains(marker, paint); else Assert.DoesNotContain(marker, paint);
                }
            }
            source = saved.Value; document = reopened;
        }
    }

    [Theory]
    [InlineData("display-type")]
    [InlineData("display-duplicate")]
    [InlineData("reference-type")]
    public void Unqualified_native_master_tokens_report_the_gap_and_reject_assignment_without_losing_source(string variant) {
        byte[] source = Variant(variant);
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, value => value.Code == "PUB_MASTER_ASSIGNMENT_UNASSESSED");
        var pages = document.Pages; var report = document.ReadReport;
        Assert.Throws<NotSupportedException>(() => document.SetPageMaster(pages[0].Id, document.MasterPages[0].Id));
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport);
        Assert.Equal(source, document.ToPubBytes(Allow));
    }

    [Theory]
    [InlineData("MultipleMasters", false)]
    [InlineData("TwoPageMaster", false)]
    [InlineData("MultipleMasters", true)]
    [InlineData("TwoPageMaster", true)]
    public void Cancelling_page_and_shape_growth_still_relocates_every_intervening_native_chunk(string fixture, bool alignment) {
        byte[] source = File.ReadAllBytes(PathFor(fixture + (alignment ? "" : "Ignored")));
        PublisherDocument document = PublisherDocument.Load(source);
        if (alignment) {
            foreach (PublisherPage page in document.Pages) document.SetPageMaster(page.Id, null);
            source = document.ToPubBytes(Allow);
            document = PublisherDocument.Load(source);
        }
        uint? alignedFrame = null;
        if (alignment) {
            for (int index = 0; index < document.Pages.Count; index++)
                document.SetPageMaster(document.Pages[index].Id, document.MasterPages[index == 2 && fixture == "MultipleMasters" ? 0 : index].Id);
            PublisherTextFrame frame = document.MasterPages[0].TextFrames[0]; alignedFrame = frame.Id;
            document.SetTextFrameLayout(frame.Id, frame.ColumnCount, frame.ColumnSpacing, frame.Padding, OfficeTextVerticalAlignment.Bottom);
        } else {
            document.SetPageMaster(document.Pages[0].Id, document.MasterPages[0].Id);
            document.SetPageMaster(document.Pages[1].Id, null);
        }
        byte[] output = document.ToPubBytes(Allow);
        Assert.Equal(Contents(source).Bytes.Length, Contents(output).Bytes.Length);
        PublisherDocument reopened = PublisherDocument.Load(output);
        Assert.Equal(document.Pages.Select(page => page.MasterPageId), reopened.Pages.Select(page => page.MasterPageId));
        Assert.Equal(document.TextStories.Select(story => (story.Id, story.Text)), reopened.TextStories.Select(story => (story.Id, story.Text)));
        if (alignedFrame.HasValue)
            Assert.Equal(OfficeTextVerticalAlignment.Bottom, reopened.MasterPages.SelectMany(page => page.TextFrames).Single(frame => frame.Id == alignedFrame).VerticalAlignment);
        AssertUnrelatedContents(source, output, alignedFrame);
        PublisherNativeWritingTests.AssertPreservedStorage(source, output, "Contents");
    }

    [Fact]
    public void Ambiguous_native_master_reference_is_rejected_at_the_read_boundary() {
        InvalidDataException failure = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(Variant("reference-duplicate")));
        Assert.Contains("0x0D", failure.Message);
    }

    [Fact]
    public void Master_suppression_composes_with_frame_alignment_and_picture_resource_relocation() {
        byte[] source = File.ReadAllBytes(PathFor("MasterContent"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherPage master = Assert.Single(document.MasterPages);
        PublisherTextFrame frame = Assert.Single(master.TextFrames, item => document.TextStories.Any(story => story.Id == item.StoryId
            && story.Text.StartsWith("Master: shared content", StringComparison.Ordinal)));
        PublisherPicture picture = Assert.Single(master.Pictures);
        PublisherPicture independent = Assert.Single(document.Pages[2].Pictures);
        document.SetPageMaster(document.Pages[0].Id, null);
        document.SetTextFrameLayout(frame.Id, frame.ColumnCount, frame.ColumnSpacing, frame.Padding, OfficeTextVerticalAlignment.Bottom);
        document.ReplacePictureImage(picture.Id, PublisherPictureEditingFixture.Image(OfficeColor.Blue));
        var saved = document.ToPubBytesResult(Allow);
        Assert.False(saved.Report.HasErrors);
        PublisherDocument reopened = PublisherDocument.Load(saved.Value);
        Assert.Null(reopened.Pages[0].MasterPageId);
        Assert.Equal(master.Id, reopened.Pages[1].MasterPageId);
        Assert.Equal(OfficeTextVerticalAlignment.Bottom, reopened.MasterPages[0].TextFrames.Single(item => item.Id == frame.Id).VerticalAlignment);
        Assert.NotEqual(reopened.MasterPages[0].Pictures[0].ImageId, reopened.Pages[2].Pictures[0].ImageId);
        Assert.Equal(independent.ImageId, reopened.Pages[2].Pictures[0].ImageId);
        Assert.Equal(2, reopened.Images.Count);
        PublisherNativeWritingTests.AssertPreservedStorage(source, saved.Value, "Contents", "Escher/EscherStm", "Escher/EscherDelayStm");
        Assert.DoesNotContain("Master: shared content", reopened.ToSvg(0));
        Assert.Contains("Master: shared content", reopened.ToSvg(1));
    }

    [Fact]
    public void Assignment_save_keeps_strict_policy_resource_cancellation_and_new_topology_guards() {
        byte[] source = File.ReadAllBytes(PathFor("MultipleMasters"));
        PublisherDocument document = PublisherDocument.Load(source);
        document.SetPageMaster(document.Pages[0].Id, document.MasterPages[1].Id);
        OfficeConversionException blocked = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        Assert.False(Assert.IsType<PublisherSaveReport>(blocked.Report).HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow, MaxOutputBytes = 1 }));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.ToPubBytes(Allow, cancellation.Token));
        Assert.Equal(document.MasterPages[1].Id, PublisherDocument.Load(document.ToPubBytes(Allow)).Pages[0].MasterPageId);
        PublisherPage created = document.AddMasterPage("No native record");
        document.SetPageMaster(document.Pages[0].Id, created.Id);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(Allow));
    }

    private static string PathFor(string name) => PublisherNativeTests.Fixture("NativeComparison/" + name + "/native-" + name.ToLowerInvariant() + ".pub");

    private static (byte[] Bytes, PublisherContents Model, PublisherBlockReader Blocks) Contents(byte[] source) {
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? compound, out string? error), error);
        byte[] bytes = compound!.Streams["Contents"];
        var data = new PublisherBinaryData(bytes, "Contents");
        var context = new PublisherParseContext(new PublisherReadOptions(), default);
        return (bytes, new PublisherContentsReader(data, context).Read(), new PublisherBlockReader(data, context));
    }

    private static void AssertUnrelatedContents(byte[] before, byte[] after, uint? alignedFrame = null) {
        var a = Contents(before); var b = Contents(after);
        Assert.Equal(a.Model.Chunks.Keys.OrderBy(id => id), b.Model.Chunks.Keys.OrderBy(id => id));
        var physical = a.Model.Pages.Where(page => !page.IsMaster).Select(page => page.Id).ToHashSet();
        foreach (var entry in a.Model.Chunks) {
            PublisherContentChunk x = entry.Value, y = b.Model.Chunks[entry.Key];
            Assert.Equal((x.Kind, x.Parent), (y.Kind, y.Parent));
            bool Retained(PublisherBlock field, uint id) => !(physical.Contains(id) && field.Id is 7 or 0x0D) && !(alignedFrame == id && field.Id == 0x35);
            var left = a.Blocks.Chunk(x.Offset, x.End).Where(field => Retained(field, x.Id)).ToArray();
            var right = b.Blocks.Chunk(y.Offset, y.End).Where(field => Retained(field, y.Id)).ToArray();
            Assert.Equal(left.Length, right.Length);
            for (int index = 0; index < left.Length; index++)
                Assert.Equal(a.Bytes.Skip(left[index].Offset - 2).Take(left[index].Length + 2), b.Bytes.Skip(right[index].Offset - 2).Take(right[index].Length + 2));
        }
    }

    private static byte[] Variant(string variant) => PublisherInputContractTests.Mutate("Contents", bytes => {
        var context = new PublisherParseContext(new PublisherReadOptions(), default);
        var data = new PublisherBinaryData(bytes, "Contents");
        PublisherContents model = new PublisherContentsReader(data, context).Read();
        PublisherContentChunk chunk = model.Chunks[model.Pages.First(page => !page.IsMaster).Id];
        var fields = new PublisherBlockReader(data, context).Chunk(chunk.Offset, chunk.End);
        PublisherBlock reference = fields.Single(field => field.Id == 0x0D);
        if (variant == "display-type") bytes[fields.Single(field => field.Id == 7).Offset - 1] = 0x0A;
        else if (variant == "display-duplicate") bytes[fields.Single(field => field.Id == 6).Offset - 2] = 7;
        else if (variant == "reference-type") bytes[reference.Offset - 1] = 0x20;
        else {
            PublisherBlock spare = fields.Single(field => field.Id == 9);
            bytes[spare.Offset - 2] = 0x0D; bytes[spare.Offset - 1] = 0x68;
            Buffer.BlockCopy(BitConverter.GetBytes(reference.Value), 0, bytes, spare.Offset, 4);
        }
    }, "NativeComparison/MultipleMastersIgnored/native-multiplemastersignored.pub");
}
