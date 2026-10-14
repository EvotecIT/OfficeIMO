using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkKeynoteParagraphEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Apple_qualified_paragraphs_and_native_resave_keep_images_and_allow_another_edit(bool nativeResave) {
        var original = IWorkSourceDocument.Open(Fixture("keynote-created-images-edited-v15.4.key"));
        var deck = original.ReadKeynote(); var edit = original.BeginEdit();
        string[] replacements = ["Edited PNG — image behind text", "Edited alpha PNG", "Edited JPEG"];
        for (int index = 0; index < 3; index++)
            edit.SetKeynoteTextBoxParagraphText(deck.Slides[index].TextBoxes[0].SourceIdentity!, 1, replacements[index]);
        Assert.Equal(File.ReadAllBytes(Fixture("keynote-preserved-paragraphs-v15.4.key")), edit.SaveBytes(AcceptCaches()));
        var source = IWorkSourceDocument.Open(Fixture("keynote-preserved-paragraphs" + (nativeResave ? "-edited" : "") + "-v15.4.key"));
        var output = source.ReadKeynote(); Assert.True(output.HasEditableContent); Assert.Empty(output.Diagnostics);
        if (nativeResave) replacements[2] = "Native checked JPEG";
        for (int index = 0; index < 3; index++) {
            Assert.Equal(replacements[index], output.Slides[index].TextBoxes[0].Content.Paragraphs[0].Text);
            var image = Assert.Single(output.Slides[index].Images); var before = Assert.Single(deck.Slides[index].Images);
            Assert.Equal(before.Geometry!.WidthPoints, image.Geometry!.WidthPoints);
            Assert.Equal(before.Geometry.HeightPoints, image.Geometry.HeightPoints);
            Assert.Equal(before.GetBytes(), image.GetBytes());
        }
        Assert.Equal(2, output.Slides[1].TextBoxes[1].Content.Paragraphs[0].Runs.Count);
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(source.BeginEdit().SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream(); converted.Value.Save(bytes);
        using var xml = PresentationDocument.Open(new MemoryStream(bytes.ToArray()), false);
        var texts = xml.PresentationPart!.SlideParts.SelectMany(part => part.Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            .Select(item => item.InnerText).ToArray();
        foreach (string expected in replacements) Assert.Contains(expected, texts);
        var next = source.BeginEdit(); var first = output.Slides[0].TextBoxes[0];
        next.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, "Native PNG — image behind text");
        Assert.Equal("Native PNG — image behind text", IWorkSourceDocument.Open(next.SaveBytes(AcceptCaches())).ReadKeynote().Slides[0].TextBoxes[0].Content.Paragraphs[0].Text);
        next.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, first.Content.Paragraphs[0].Text);
        Assert.Equal(source.ComputePackageContentHash(), next.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Paragraph_replacements_compose_across_shapes_preserve_other_bytes_and_save_complete_PPTX() {
        var source = Source(); var deck = source.ReadKeynote(); var edit = source.BeginEdit();
        var first = deck.Slides[0].TextBoxes[0]; var other = deck.Slides[1].TextBoxes[0];
        edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, "Bravo café 😃");
        edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, "Edited body");
        edit.SetKeynoteTextBoxParagraphText(other.SourceIdentity!, 1, "Other slide");
        var result = edit.SaveBytesResult(AcceptCaches()); var saved = IWorkSourceDocument.Open(result.GetPackageBytes());
        var output = saved.ReadKeynote(); Assert.True(output.HasEditableContent); Assert.Empty(output.Diagnostics);
        Assert.Equal("Bravo café 😃\nEdited body\n", output.Slides[0].TextBoxes[0].Content.PlainText);
        Assert.Equal("Other slide\n", output.Slides[1].TextBoxes[0].Content.PlainText);
        Assert.Equal("Alpha café 😀\nSecond body\n", first.Content.PlainText);
        Assert.Equal(2, result.Report.UpdatedRecords.Count);
        Assert.Equal(deck.Slides[0].Drawables.Select(item => item.TextBox!.SourceIdentity!.RecordIdentifier),
            output.Slides[0].Drawables.Select(item => item.TextBox!.SourceIdentity!.RecordIdentifier));
        foreach (var record in source.Records) {
            var after = saved.Records.Single(item => item.Identifier == record.Identifier && item.PayloadIndex == record.PayloadIndex);
            Assert.Equal(record.PayloadLength, after.PayloadLength); Assert.Equal(record.ObjectReferences, after.ObjectReferences);
            Assert.Equal(record.Versions, after.Versions); Assert.Equal(record.DataReferences, after.DataReferences);
            if (!result.Report.UpdatedRecords.Any(item => item.RecordIdentifier == record.Identifier)) Assert.Equal(record.GetPayload(), after.GetPayload());
        }
        foreach (var entry in source.Entries.Where(item => !result.Report.UpdatedEntryPaths.Contains(item.Path)))
            Assert.Equal(entry.GetBytes(), saved.Entries.Single(item => item.Path == entry.Path).GetBytes());
        Assert.Equal(first.Geometry!.LeftPoints, output.Slides[0].TextBoxes[0].Geometry!.LeftPoints);
        Assert.Equal(first.Geometry.WidthPoints, output.Slides[0].TextBoxes[0].Geometry!.WidthPoints);
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(result.GetPackageBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream(); converted.Value.Save(bytes);
        using var xml = PresentationDocument.Open(new MemoryStream(bytes.ToArray()), false);
        var texts = xml.PresentationPart!.SlideParts.SelectMany(part => part.Slide!.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
            .Select(item => item.InnerText).ToArray();
        foreach (string expected in new[] { "Bravo café 😃", "Edited body", "Other slide", "Rich tail" }) Assert.Contains(expected, texts);
    }

    [Fact]
    public void Rejected_and_cancelled_replacements_preserve_pending_state_and_restoring_text_removes_edits() {
        var source = Source(); var first = source.ReadKeynote().Slides[0].TextBoxes[0]; var edit = source.BeginEdit();
        edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, "Bravo café 😃");
        edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, "Edited body");
        byte[] before = edit.SaveBytes(AcceptCaches());
        Assert.Throws<NotSupportedException>(() => edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, "short"));
        Assert.Throws<NotSupportedException>(() => edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, "Edited bódy"));
        Assert.Throws<ArgumentException>(() => edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, "Edited\nbody"));
        Assert.Throws<ArgumentException>(() => edit.SetKeynoteTextBoxParagraphText(first.Content.SourceIdentity!, 2, "Edited body"));
        Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 0, ""));
        Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 99, ""));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, "Second body", cancellation.Token));
        Assert.Equal(before, edit.SaveBytes(AcceptCaches()));
        using var stream = new MemoryStream(); stream.WriteByte(9);
        Assert.Throws<InvalidOperationException>(() => edit.Save(stream)); Assert.Equal(new byte[] { 9 }, stream.ToArray());
        edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 1, first.Content.Paragraphs[0].Text);
        Assert.Single(edit.SaveBytesResult(AcceptCaches()).Report.UpdatedRecords);
        edit.SetKeynoteTextBoxParagraphText(first.SourceIdentity!, 2, first.Content.Paragraphs[1].Text);
        Assert.Empty(edit.SaveBytesResult().Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Selected_mixed_runs_and_shared_storage_are_rejected_before_mutation() {
        var original = Source(); var boxes = original.ReadKeynote().Slides[0].TextBoxes; var edit = original.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetKeynoteTextBoxParagraphText(boxes[1].SourceIdentity!, 1, "Other end"));
        Assert.Equal(original.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
        var shape = original.Records.Single(item => item.Identifier == boxes[1].SourceIdentity!.RecordIdentifier && item.PayloadIndex == 0);
        byte[] payload = shape.GetPayload();
        byte[] reference = new IWorkProtoWriter(32, default).UInt(1, boxes[0].Content.SourceIdentity!.RecordIdentifier).ToArray();
        foreach (int field in new[] { 2, 4 }) IWorkPreservingWireEdit.ReplaceBytesRange(payload, new[] { (field, -1) }, 0, reference, new(), default);
        var package = IWorkPreservationPackageWriter.Write(original, new Dictionary<IWorkArchiveRecord, byte[]> { [shape] = payload }, AcceptCaches(), default);
        var shared = IWorkSourceDocument.Open(package.GetPackageBytes()); Assert.True(shared.ReadKeynote().HasEditableContent);
        edit = shared.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetKeynoteTextBoxParagraphText(boxes[0].SourceIdentity!, 1, "Bravo café 😃"));
        Assert.Equal(shared.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Locked_shape_is_rejected_without_mutating_its_native_text() {
        var original = IWorkSourceDocument.Open(Fixture("keynote-created-images-edited-v15.4.key"));
        var box = original.ReadKeynote().Slides[0].TextBoxes[0];
        var record = original.Records.Single(item => item.Identifier == box.SourceIdentity!.RecordIdentifier && item.PayloadIndex == 0);
        var root = IWorkProtobuf.Parse(record.GetPayload(), new());
        var shape = IWorkProtobuf.Parse(root.GetBytes(1)!, new());
        byte[] drawable = shape.GetBytes(1)!; bool found = false;
        // Flip the Apple fixture's existing false lock without changing lengths
        // or reserializing unrelated declarations and references.
        for (int offset = 0; offset < drawable.Length;) {
            ulong tag = IWorkProtobuf.ReadVarint(drawable, ref offset);
            if (tag == 40) { Assert.False(found); Assert.Equal(0ul, IWorkProtobuf.ReadVarint(drawable, ref offset)); drawable[offset - 1] = 1; found = true; }
            else switch (tag & 7) {
                case 0: _ = IWorkProtobuf.ReadVarint(drawable, ref offset); break;
                case 1: offset += 8; break;
                case 2: int length = checked((int)IWorkProtobuf.ReadVarint(drawable, ref offset)); offset += length; break;
                case 5: offset += 4; break;
                default: throw new InvalidDataException("Unexpected fixture wire kind.");
            }
        }
        Assert.True(found); byte[] payload = record.GetPayload();
        IWorkPreservingWireEdit.ReplaceBytesRange(payload, new[] { (1, -1), (1, -1) }, 0, drawable, new(), default);
        var package = IWorkPreservationPackageWriter.Write(original, new Dictionary<IWorkArchiveRecord, byte[]> { [record] = payload }, AcceptCaches(), default);
        var locked = IWorkSourceDocument.Open(package.GetPackageBytes()); Assert.True(locked.ReadKeynote().HasEditableContent);
        var edit = locked.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetKeynoteTextBoxParagraphText(box.SourceIdentity!, 1, "Edited PNG — image behind text"));
        Assert.Equal(locked.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Nested_index_edits_are_rejected_without_rewriting_the_container() {
        var original = Source(); using var inner = new MemoryStream();
        using (var index = new ZipArchive(inner, ZipArchiveMode.Create, true))
            foreach (var entry in original.Entries.Where(item => item.Path.StartsWith("Index/", StringComparison.Ordinal)))
                using (var stream = index.CreateEntry(entry.Path).Open()) { stream.Write(entry.GetBytes()); }
        using var outer = new MemoryStream();
        using (var package = new ZipArchive(outer, ZipArchiveMode.Create, true)) {
            using (var stream = package.CreateEntry("Index.zip").Open()) { stream.Write(inner.ToArray()); }
            foreach (var entry in original.Entries.Where(item => !item.Path.StartsWith("Index/", StringComparison.Ordinal)))
                using (var stream = package.CreateEntry(entry.Path).Open()) { stream.Write(entry.GetBytes()); }
        }
        var nested = IWorkSourceDocument.Open(outer.ToArray(), IWorkDocumentKind.Keynote); var edit = nested.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetKeynoteTextBoxParagraphText(nested.ReadKeynote().Slides[0].TextBoxes[0].SourceIdentity!, 1, "Bravo café 😃"));
        Assert.Equal(nested.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    private static IWorkSourceDocument Source() {
        var model = IWorkKeynoteDocument.Create(); var slide = model.AddSlide();
        slide.AddText("Alpha café 😀\nSecond body", 40, 40, 600, 150);
        slide.AddRichText(new[] { new IWorkTextRunData("Rich", bold: true), new IWorkTextRunData(" tail") }, 40, 240, 600, 100);
        model.AddSlide().AddText("Final slide", 60, 60, 600, 120);
        return IWorkSourceDocument.Open(model.SaveBytes());
    }
    private static IWorkNativeSaveOptions AcceptCaches() => new() { AllowUnassessedVisualCaches = true };
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
