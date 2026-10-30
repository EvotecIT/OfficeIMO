using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeWritingTests {
    private const long MaximumBytes = 64L * 1024 * 1024;

    [Fact]
    public void Unchanged_native_content_and_semantic_noops_return_exact_original_bytes_with_owned_results() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("native-authored.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.True(prepared.IsPrepared); Assert.False(prepared.HasLoss); prepared.RequireNoLoss();
        Assert.Equal(source, prepared.CopyBytes());
        byte[] detached = prepared.CopyBytes(); detached[0] ^= 0xFF;
        Assert.Equal(source, prepared.CopyBytes());

        byte[] text = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Text/native-text.pub"));
        PublisherDocument unchanged = PublisherDocument.Load(text);
        PublisherTextStory story = Assert.Single(unchanged.TextStories);
        unchanged.ReplaceStoryParagraphs(story.Id, story.Paragraphs);
        PublisherTextFrame frame = Assert.Single(unchanged.Pages[0].TextFrames);
        unchanged.SetTextFrameBounds(frame.Id, frame.X, frame.Y, frame.Width, frame.Height);
        unchanged.SetTextFrameLayout(frame.Id, frame.ColumnCount, frame.ColumnSpacing, frame.Padding, frame.VerticalAlignment);
        Assert.Equal(text, unchanged.PrepareNativeBytes(MaximumBytes).CopyBytes());
    }

    [Fact]
    public void Rectangle_bounds_and_existing_fill_slot_change_one_native_object_and_preserve_other_streams_and_metadata() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherShape original = Assert.Single(document.Pages[0].Shapes);
        OfficeShape paint = original.CreateShape();
        document.SetShapeBounds(original.Id, 36, 48, 120, 84);
        document.SetShapePaint(original.Id, OfficeColor.Blue, paint.StrokeColor, paint.StrokeWidth,
            paint.FillOpacity ?? 1, paint.StrokeOpacity ?? 1);
        var pages = document.Pages; var report = document.SourceReadReport;
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.True(prepared.IsPrepared, string.Join("; ", prepared.FidelityDiagnostics.Select(item => item.Message)));
        Assert.Contains(prepared.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_FIXED_PATCH_UNQUALIFIED"
            && item.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Throws<OfficeConversionException>(prepared.RequireNoLoss);
        byte[] output = prepared.CopyBytes();
        AssertPreservedStorage(source, output);
        OfficeCompoundFile before = Compound(source), after = Compound(output);
        PublisherEscherData raw = Artwork(before);
        PublisherNativeShapeBinding binding = raw.Shapes.Single(shape => shape.Id == original.Id).NativeBinding!;
        // These offsets and identities come from the independent raw control
        // inspection of the existing native-rectangle.pub fixture.
        Assert.Equal(293U, binding.Id); Assert.Equal(1025U, binding.Spid);
        Assert.Equal(new[] { 1534, 1540, 1546, 1552 }, binding.Anchor.Select(slot => slot.Offset));
        var permitted = new HashSet<int>(new[] { 1258, 1534, 1540, 1546, 1552 }.SelectMany(offset => Enumerable.Range(offset, 4)));
        PublisherNativePropertySlot color = Assert.Single(binding.Properties, slot => slot.Id == 0x0181);
        Assert.Equal(1258, color.Value.Offset); Assert.Equal((ushort)0x0181, color.RawOperationId);
        int[] changed = before.Streams["Escher/EscherStm"].Select((value, index) => (value, index))
            .Where(item => item.value != after.Streams["Escher/EscherStm"][item.index]).Select(item => item.index).ToArray();
        Assert.NotEmpty(changed); Assert.All(changed, offset => Assert.Contains(offset, permitted));
        PublisherDocument reopened = PublisherDocument.Load(output);
        PublisherShape shape = Assert.Single(reopened.Pages[0].Shapes);
        PublisherTextFrame frame = Assert.Single(reopened.Pages[0].TextFrames);
        Assert.Equal(original.Id, shape.Id);
        Assert.Equal((36D, 48D, 120D, 84D), (shape.X, shape.Y, shape.Width, shape.Height));
        Assert.Equal((shape.X, shape.Y, shape.Width, shape.Height), (frame.X, frame.Y, frame.Width, frame.Height));
        Assert.Equal(OfficeColor.Blue, shape.CreateShape().FillColor);
        PublisherNativeShapeBinding saved = Artwork(after).Shapes.Single(item => item.Id == shape.Id).NativeBinding!;
        Assert.Equal(binding.Spid, saved.Spid); Assert.Equal(binding.BodyOffset, saved.BodyOffset); Assert.Equal(binding.BodyLength, saved.BodyLength);
        Assert.Same(pages, document.Pages); Assert.Same(report, document.SourceReadReport);
    }

    [Fact]
    public void Centered_and_paired_master_anchors_use_their_native_registration_and_keep_other_content() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TwoPageMaster/native-twopagemaster.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextFrame[] originals = document.MasterPages.SelectMany(page => page.TextFrames)
            .Concat(document.Pages[0].TextFrames).ToArray();
        Assert.Equal(4, originals.Length);
        foreach (PublisherTextFrame frame in originals)
            document.SetTextFrameBounds(frame.Id, frame.X + 6, frame.Y + 12, frame.Width, frame.Height);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.True(prepared.IsPrepared, string.Join("; ", prepared.FidelityDiagnostics.Select(item => item.Message)));
        byte[] output = prepared.CopyBytes(); AssertPreservedStorage(source, output);
        PublisherDocument reopened = PublisherDocument.Load(output);
        var frames = reopened.MasterPages.Concat(reopened.Pages).SelectMany(page => page.TextFrames).ToDictionary(frame => frame.Id);
        foreach (PublisherTextFrame frame in originals) {
            Assert.Equal(frame.X + 6, frames[frame.Id].X, 6); Assert.Equal(frame.Y + 12, frames[frame.Id].Y, 6);
            Assert.Equal(frame.Width, frames[frame.Id].Width, 6); Assert.Equal(frame.Height, frames[frame.Id].Height, 6);
        }
        Assert.Equal(document.TextStories.Select(story => story.Text), reopened.TextStories.Select(story => story.Text));
    }

    [Fact]
    public void Picture_bounds_keep_shared_asset_bindings_and_other_placements_unchanged() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/PictureOwnership/native-pictureownership.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherPicture original = document.Pages[0].Pictures[0], sibling = document.Pages[0].Pictures[1];
        document.SetPictureBounds(original.Id, 84, 168, 120, 72);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.True(prepared.IsPrepared, string.Join("; ", prepared.FidelityDiagnostics.Select(item => item.Message)));
        byte[] output = prepared.CopyBytes(); AssertPreservedStorage(source, output);
        PublisherDocument reopened = PublisherDocument.Load(output);
        PublisherPicture current = reopened.Pages[0].Pictures.Single(picture => picture.Id == original.Id);
        PublisherPicture retained = reopened.Pages[0].Pictures.Single(picture => picture.Id == sibling.Id);
        Assert.Equal((84D, 168D, 120D, 72D), (current.X, current.Y, current.Width, current.Height));
        Assert.Equal(original.ImageId, current.ImageId); Assert.Equal(sibling.ImageId, retained.ImageId);
        Assert.Equal((sibling.X, sibling.Y, sibling.Width, sibling.Height), (retained.X, retained.Y, retained.Width, retained.Height));
        Assert.Equal(document.Images.Select(image => image.Id), reopened.Images.Select(image => image.Id));
        foreach (PublisherImage image in document.Images) Assert.Equal(image.GetBytes(), reopened.Images.Single(item => item.Id == image.Id).GetBytes());
    }

    [Theory]
    [InlineData("missing-opacity")]
    [InlineData("visibility")]
    [InlineData("alpha")]
    public void Paint_changes_outside_existing_scalar_slots_refuse_the_whole_candidate(string change) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
        PublisherShape shape = document.Pages[0].Shapes[0]; OfficeShape prior = shape.CreateShape();
        OfficeColor? fill = change == "visibility" ? null : change == "alpha" ? OfficeColor.FromRgba(255, 0, 0, 128) : prior.FillColor;
        document.SetShapePaint(shape.Id, fill, prior.StrokeColor, prior.StrokeWidth,
            change == "missing-opacity" ? 0.5 : prior.FillOpacity ?? 1, prior.StrokeOpacity ?? 1);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.False(prepared.IsPrepared);
        Assert.Contains(prepared.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Failure
            && item.Message.Contains(change == "missing-opacity" ? "0x0182" : "visibility or alpha"));
    }

    [Theory]
    [InlineData("multi-paragraph-story")]
    [InlineData("format")]
    [InlineData("page")]
    [InlineData("master")]
    [InlineData("assignment")]
    [InlineData("continuation")]
    [InlineData("text-frame")]
    [InlineData("frame-layout")]
    [InlineData("shape")]
    [InlineData("picture")]
    [InlineData("image-replacement")]
    [InlineData("table-cell")]
    [InlineData("table-bounds")]
    [InlineData("table")]
    [InlineData("source-free")]
    public void Unsupported_current_mutation_families_refuse_output_instead_of_copying_the_unchanged_source(string mutation) {
        string file = mutation.StartsWith("table-", StringComparison.Ordinal) ? "NativeComparison/TableLabel/native-tablelabel.pub"
            : mutation == "image-replacement" ? "NativeComparison/Image/native-image.pub" : "NativeComparison/Text/native-text.pub";
        PublisherDocument document = mutation == "source-free" ? PublisherDocument.Create(612, 792)
            : PublisherDocument.Load(PublisherNativeTests.Fixture(file));
        PublisherPage page = document.Pages[0];
        switch (mutation) {
            case "multi-paragraph-story": document.ReplaceStoryParagraphs(document.TextStories[0].Id,
                Paragraphs("First paragraph").Concat(Paragraphs("Second paragraph")).ToArray()); break;
            case "format":
                OfficeRichTextRun run = document.TextStories[0].Paragraphs[0].Runs[0];
                document.ReplaceStoryParagraphs(document.TextStories[0].Id, new[] { new OfficeRichTextParagraph(new[] {
                    new OfficeRichTextRun(run.Text, run.FontSize, run.Color, !run.Bold, run.Italic, run.Underline, run.FontFamily)
                }) }); break;
            case "page": document.AddPage(); break;
            case "master": document.AddMasterPage("New master"); break;
            case "assignment": document.SetPageMaster(page.Id, null); break;
            case "continuation": document.AddLinkedTextFrame(page.TextFrames[0].Id, page.Id, 40, 300, 200, 60); break;
            case "text-frame": document.AddTextFrame(page.Id, 40, 300, 200, 60, Paragraphs("New story")); break;
            case "frame-layout": document.SetTextFrameLayout(page.TextFrames[0].Id, 2, 12, new OfficeTextPadding(2, 2, 2, 2), OfficeTextVerticalAlignment.Center); break;
            case "shape": document.AddShape(page.Id, OfficeShape.Rectangle(40, 30), 40, 300); break;
            case "picture": document.AddPicture(page.Id, PictureBytes(), 40, 300, 40, 30); break;
            case "image-replacement": document.ReplacePictureImage(page.Pictures[0].Id, PictureBytes()); break;
            case "table-cell": document.ReplaceTableCellParagraphs(page.Tables[0].Id, 0, 0, Paragraphs("Changed cell")); break;
            case "table-bounds":
                PublisherTable table = page.Tables[0]; document.SetTableBounds(table.Id, table.X + 6, table.Y, table.Width, table.Height); break;
            case "table": document.AddTable(page.Id, 40, 300, new[] { 60D }, new[] { 30D }); break;
        }
        var pages = document.Pages; var stories = document.TextStories; var images = document.Images;
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.False(prepared.IsPrepared); Assert.Equal(0, prepared.ByteLength); Assert.True(prepared.HasLoss);
        Assert.Contains(prepared.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Failure);
        if (mutation == "frame-layout")
            Assert.Contains(prepared.FidelityDiagnostics, item => item.Message.Contains("frame input other than its bounds"));
        Assert.Throws<NotSupportedException>(() => prepared.CopyBytes());
        Assert.Same(pages, document.Pages); Assert.Same(stories, document.TextStories); Assert.Same(images, document.Images);
    }

    [Fact]
    public void Nonrepresentable_anchor_and_opaque_active_carrier_refuse_edited_output_but_leave_exact_copy_available() {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
        PublisherDocument fractional = PublisherDocument.Load(source);
        PublisherShape shape = fractional.Pages[0].Shapes[0];
        fractional.SetShapeBounds(shape.Id, 1D / 3, 48, 120, 84);
        Assert.False(fractional.PrepareNativeBytes(MaximumBytes).IsPrepared);

        byte[] carrier = OfficeCompoundFileWriter.Rewrite(Compound(source), new Dictionary<string, byte[]> { ["Objects/opaque"] = new byte[] { 1, 2, 3 } });
        PublisherDocument active = PublisherDocument.Load(carrier);
        Assert.Equal(carrier, active.PrepareNativeBytes(MaximumBytes).CopyBytes());
        active.SetShapeBounds(active.Pages[0].Shapes[0].Id, 36, 48, 120, 84);
        PublisherNativePreparation refused = active.PrepareNativeBytes(MaximumBytes);
        Assert.False(refused.IsPrepared);
        Assert.Contains(refused.FidelityDiagnostics, item => item.Message.Contains("active-content"));
        Assert.Throws<OperationCanceledException>(() => active.PrepareNativeBytes(MaximumBytes, new CancellationToken(true)));
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(source).PrepareNativeBytes(source.Length - 1));
    }

    [Fact]
    public void Unreachable_directory_storage_keeps_exact_copy_available_but_refuses_edited_output() {
        OfficeCompoundFile original = Compound(File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub")));
        var entries = original.Entries.Where(entry => !entry.IsFallback).Concat(new[] {
            new OfficeCompoundFileEntry("Detached", "Proof/Detached", 1, 0)
        }).ToArray();
        var fixture = new OfficeCompoundFile(original.Streams, entries, original.RootEntry);
        var streams = fixture.Streams.Select(pair => new OfficeCompoundStream(pair.Key, pair.Value)).ToArray();
        OfficeCompoundWriterLayout layout = OfficeCompoundWriterLayout.Create(streams, fixture);
        OfficeCompoundWriterEntry parent = layout.Entries.Single(entry => entry.Path == "Proof");
        byte[] bytes = OfficeCompoundFileWriter.Rewrite(fixture, new Dictionary<string, byte[]>(), maxOutputBytes: MaximumBytes);
        // Core's version-3 writer emits its directory contiguously. Removing
        // this sole child leaves a real active entry outside the reachable tree.
        int directory = checked(((int)BitConverter.ToUInt32(bytes, 48) + 1) * 512);
        int childLink = checked(directory + parent.DirectoryIndex * 128 + 76);
        Assert.NotEqual(uint.MaxValue, BitConverter.ToUInt32(bytes, childLink));
        BitConverter.GetBytes(uint.MaxValue).CopyTo(bytes, childLink);
        using (var input = new MemoryStream(bytes, writable: false))
        using (var index = new OfficeCompoundFileReader.ReadSession(input,
            new OfficeCompoundReadOptions(maxStreamBytes: MaximumBytes, maxTotalStreamBytes: MaximumBytes), leaveOpen: true, CancellationToken.None)) {
            Assert.True(index.HasCanonicalRootDirectory); Assert.True(index.HasInvalidDirectoryTree);
        }
        PublisherDocument document = PublisherDocument.Load(bytes);
        Assert.Equal(bytes, document.PrepareNativeBytes(MaximumBytes).CopyBytes());
        document.SetShapeBounds(document.Pages[0].Shapes[0].Id, 36, 48, 120, 84);
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.False(prepared.IsPrepared);
        Assert.Contains(prepared.FidelityDiagnostics, item => item.LossKind == OfficeConversionLossKind.Failure
            && item.Message.Contains("invalid directory tree"));
    }

    [Theory]
    [InlineData("Text", "Publisher text with a longer replacement", 19456, 19968, false)]
    [InlineData("TextLong", "Publisher text", 18944, 19456, false)]
    [InlineData("Text", "Changed string", 18944, 19456, false)]
    [InlineData("Text", "Text and frame changed together", 19456, 19968, true)]
    public void Plain_story_text_replacement_rewrites_native_ranges_and_preserves_other_storage(
        string fixture, string replacement, int expectedParagraphPage, int expectedCharacterPage, bool moveFrame) {
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture($"NativeComparison/{fixture}/native-{fixture.ToLowerInvariant()}.pub"));
        PublisherDocument document = PublisherDocument.Load(source);
        PublisherTextStory story = Assert.Single(document.TextStories);
        OfficeRichTextParagraph paragraph = Assert.Single(story.Paragraphs);
        OfficeRichTextRun run = Assert.Single(paragraph.Runs);
        document.ReplaceStoryParagraphs(story.Id, new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(replacement, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily)
        }, paragraph.Alignment, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor).WithTabStops(paragraph.TabStops) });
        if (moveFrame) {
            PublisherTextFrame frame = document.Pages[0].TextFrames[0];
            document.SetTextFrameBounds(frame.Id, frame.X + 6, frame.Y + 12, frame.Width, frame.Height);
        }
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.True(prepared.IsPrepared, string.Join("; ", prepared.FidelityDiagnostics.Select(item => item.Message)));
        Assert.Contains(prepared.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_QUILL_TEXT_UNQUALIFIED"
            && item.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Throws<OfficeConversionException>(() => prepared.RequireNoLoss());
        Assert.Equal(story.Id, Assert.Single(prepared.StyleDependencies).StoryId);
        byte[] output = prepared.CopyBytes();
        AssertPreservedStorage(source, output, "Quill/QuillSub/CONTENTS", moveFrame ? new[] { "Escher/EscherStm" } : Array.Empty<string>());
        PublisherDocument reopened = PublisherDocument.Load(output);
        PublisherTextStory current = Assert.Single(reopened.TextStories);
        Assert.Equal(replacement + "\n", current.Text);
        OfficeRichTextRun actual = Assert.Single(Assert.Single(current.Paragraphs).Runs);
        Assert.Equal((run.FontFamily, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline),
            (actual.FontFamily, actual.FontSize, actual.Color, actual.Bold, actual.Italic, actual.Underline));
        Assert.Equal(document.Pages[0].TextFrames[0].X, reopened.Pages[0].TextFrames[0].X);
        Assert.Equal(document.Pages[0].TextFrames[0].Y, reopened.Pages[0].TextFrames[0].Y);
        if (moveFrame) {
            Assert.Equal((78D, 84D), (reopened.Pages[0].TextFrames[0].X, reopened.Pages[0].TextFrames[0].Y));
            Assert.Contains(prepared.FidelityDiagnostics, item => item.Code == "PUB_NATIVE_FIXED_PATCH_UNQUALIFIED");
        }
        byte[] q = Compound(output).Streams["Quill/QuillSub/CONTENTS"];
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var chunks = PublisherQuillDirectory.Read(new PublisherBinaryData(q, "Quill"), context);
        PublisherQuillChunk Find(string tag) => Assert.Single(chunks, item => item.Tag == tag);
        // The native Text/TextLong controls independently establish the TEXT
        // origin and FDPP/FDPC endpoints, including the final CR's own range.
        Assert.Equal(512, Find("TEXT").Offset);
        Assert.Equal(System.Text.Encoding.Unicode.GetBytes(replacement + "\r"),
            q.Skip(512).Take((replacement.Length + 1) * 2).ToArray());
        Assert.Equal((expectedParagraphPage, expectedCharacterPage), (Find("FDPP").Offset, Find("FDPC").Offset));
        uint contentEnd = (uint)(512 + replacement.Length * 2), finalEnd = contentEnd + 2;
        Assert.Equal(finalEnd, BitConverter.ToUInt32(q, Find("FDPP").Offset + 8));
        Assert.Equal(contentEnd, BitConverter.ToUInt32(q, Find("FDPC").Offset + 8));
        Assert.Equal(finalEnd, BitConverter.ToUInt32(q, Find("FDPC").Offset + 12));
        Assert.Equal((uint)(replacement.Length + 1), BitConverter.ToUInt32(q, Find("STRS").Offset + 12));
        Assert.Equal(finalEnd, BitConverter.ToUInt32(q, Find("BTEP").Offset + 16));
        Assert.Equal((uint)expectedParagraphPage, BitConverter.ToUInt32(q, Find("BTEP").Offset + 20));
        Assert.Equal(finalEnd, BitConverter.ToUInt32(q, Find("BTEC").Offset + 16));
        Assert.Equal((uint)expectedCharacterPage, BitConverter.ToUInt32(q, Find("BTEC").Offset + 20));
        if (fixture == "TextLong") Assert.All(q.Skip((int)finalEnd).Take(82 - (replacement.Length + 1) * 2), value => Assert.Equal(0, value));
        byte[] old = Compound(source).Streams["Quill/QuillSub/CONTENTS"];
        var oldChunks = PublisherQuillDirectory.Read(new PublisherBinaryData(old, "Quill"), context);
        foreach (PublisherQuillChunk chunk in chunks.Where(item => item.Tag is "STSH" or "FONT" or "MCLD" or "PL  " or "SGP " or "INK ")) {
            PublisherQuillChunk original = oldChunks.Single(item => item.Tag == chunk.Tag && item.Id == chunk.Id);
            Assert.Equal(old.Skip(original.Offset).Take(original.Length), q.Skip(chunk.Offset).Take(chunk.Length));
        }
        Assert.Equal(source, PublisherDocument.Load(source).PrepareNativeBytes(MaximumBytes).CopyBytes());
        Assert.Throws<InvalidDataException>(() => document.PrepareNativeBytes(1));
    }

    [Fact]
    public void Contradictory_native_formatting_index_refuses_text_replacement_but_allows_exact_copy() {
        byte[] source = PublisherInputContractTests.Mutate("Quill/QuillSub/CONTENTS", bytes => {
            // The original Text control's BTEC is [20018,20042), with the
            // character formatting page's absolute position at its final DWORD.
            Assert.Equal(19456U, BitConverter.ToUInt32(bytes, 20038));
            PublisherInputContractTests.WriteUInt32(bytes, 20038, 18944);
        }, "NativeComparison/Text/native-text.pub");
        PublisherDocument document = PublisherDocument.Load(source);
        Assert.Equal(source, document.PrepareNativeBytes(MaximumBytes).CopyBytes());
        document.ReplaceStoryParagraphs(document.TextStories[0].Id, Paragraphs("Replacement text"));
        PublisherNativePreparation prepared = document.PrepareNativeBytes(MaximumBytes);
        Assert.False(prepared.IsPrepared);
        Assert.Throws<NotSupportedException>(() => prepared.CopyBytes());
        Assert.Contains(prepared.FidelityDiagnostics, item => item.Location == "Quill" && item.LossKind == OfficeConversionLossKind.Failure);
        Assert.Equal(2U, Assert.Single(prepared.StyleDependencies).StoryId);
    }

    private static IReadOnlyList<OfficeRichTextParagraph> Paragraphs(string text) =>
        new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, 12, OfficeColor.Black, fontFamily: "Arial") }) };
    private static byte[] PictureBytes() => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.png"));
    private static OfficeCompoundFile Compound(byte[] bytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error); return compound!;
    }
    private static PublisherEscherData Artwork(OfficeCompoundFile compound) => new PublisherEscherReader(
        new PublisherBinaryData(compound.Streams["Escher/EscherStm"], "Escher/EscherStm"),
        compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay) ? new PublisherBinaryData(delay, "Escher/EscherDelayStm") : null,
        new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Read();
    private static void AssertPreservedStorage(byte[] source, byte[] output, string changedStream = "Escher/EscherStm", params string[] additionalChangedStreams) {
        OfficeCompoundFile before = Compound(source), after = Compound(output);
        Assert.Equal(before.Streams.Keys.OrderBy(path => path), after.Streams.Keys.OrderBy(path => path));
        foreach (var stream in before.Streams.Where(item => item.Key != changedStream && !additionalChangedStreams.Contains(item.Key)))
            Assert.Equal(stream.Value, after.Streams[stream.Key]);
        var metadata = before.Entries.ToDictionary(entry => entry.Path);
        Assert.Equal(metadata.Count, after.Entries.Count);
        foreach (OfficeCompoundFileEntry entry in after.Entries) {
            OfficeCompoundFileEntry original = metadata[entry.Path];
            Assert.Equal((original.ObjectType, original.ClassId, original.StateBits, original.CreationTime, original.ModifiedTime),
                (entry.ObjectType, entry.ClassId, entry.StateBits, entry.CreationTime, entry.ModifiedTime));
        }
    }
}
