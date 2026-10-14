using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork.Internal;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesParagraphEditingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Apple_qualified_edits_and_native_resave_retain_selected_content_and_allow_another_edit(bool nativeResave) {
        var original = IWorkSourceDocument.Open(Fixture("pages-created-layout-edited-v15.4.pages"));
        var edit = original.BeginEdit(); var id = original.ReadPages().Body.SourceIdentity!;
        edit.SetPagesBodyParagraphText(id, 1, "Edited paragraph layout");
        edit.SetPagesBodyParagraphText(id, 3, "Start\tLeft\tCenter\tRight\t234.56");
        edit.SetPagesBodyParagraphText(id, 7, "Updated reset Native layout edit.");
        Assert.Equal(File.ReadAllBytes(Fixture("pages-preserved-paragraphs-v15.4.pages")), edit.SaveBytes(AcceptCaches()));
        var source = IWorkSourceDocument.Open(Fixture("pages-preserved-paragraphs" + (nativeResave ? "-edited" : "") + "-v15.4.pages"));
        var pages = source.ReadPages();
        Assert.True(pages.HasEditableContent); Assert.True(pages.Body.IsComplete); Assert.Empty(pages.Diagnostics);
        Assert.Equal(8, pages.Body.Paragraphs.Count);
        Assert.Equal("Edited paragraph layout", pages.Paragraphs[0]);
        Assert.Equal("Start\tLeft\tCenter\tRight\t234.56", pages.Paragraphs[2]);
        Assert.Equal(nativeResave ? "Native checked paragraph." : "Updated reset Native layout edit.", pages.Paragraphs[6]);
        Assert.Equal("Rich layout plain Zażółć 😀", pages.Paragraphs[5]);
        Assert.Equal(2, pages.Body.Paragraphs[5].Runs.Count);
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(source.BeginEdit().SaveBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream(); converted.Value.Save(bytes); bytes.Position = 0;
        using var saved = WordDocument.Load(bytes); Assert.Empty(saved.ValidateDocument());
        using var xml = WordprocessingDocument.Open(new MemoryStream(bytes.ToArray()), false);
        Assert.Equal(pages.Paragraphs, xml.MainDocumentPart!.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
            .Select(item => string.Concat(item.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().SelectMany(run => run.ChildElements)
                .Select(node => node switch { DocumentFormat.OpenXml.Wordprocessing.TabChar => "\t",
                    DocumentFormat.OpenXml.Wordprocessing.Text text => text.Text, _ => "" }))));
        var next = source.BeginEdit();
        next.SetPagesBodyParagraphText(pages.Body.SourceIdentity!, 1, "Native paragraph layout");
        Assert.Equal("Native paragraph layout", IWorkSourceDocument.Open(next.SaveBytes(AcceptCaches())).ReadPages().Paragraphs[0]);
        next.SetPagesBodyParagraphText(pages.Body.SourceIdentity!, 1, pages.Paragraphs[0]);
        Assert.Equal(source.ComputePackageContentHash(), next.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void A_source_with_nonempty_header_content_is_rejected_before_mutation() {
        var original = Source(); var header = original.ReadPages().HeaderContents.First();
        var record = original.Records.Single(item => item.Identifier == header.SourceIdentity!.RecordIdentifier && item.PayloadIndex == 0);
        byte[] payload = record.GetPayload();
        IWorkPreservingWireEdit.ReplaceBytesRange(payload, new[] { (3, -1) }, 0, new byte[] { (byte)'x' }, new(), default);
        var populated = IWorkPreservationPackageWriter.Write(original, new Dictionary<IWorkArchiveRecord, byte[]> { [record] = payload }, AcceptCaches(), default);
        var source = IWorkSourceDocument.Open(populated.GetPackageBytes()); var pages = source.ReadPages();
        Assert.True(pages.HasEditableContent);
        Assert.Contains(pages.HeaderContents.SelectMany(item => item.Paragraphs), item => item.Text == "x");
        var edit = source.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(pages.Body.SourceIdentity!, 2, "Edited text"));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Paragraph_edits_compose_preserve_native_structure_and_survive_saved_Word_readback() {
        var source = Source(); var body = source.ReadPages().Body; var edit = source.BeginEdit();
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 1, "Bravo café 😃");
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 2, "Edited text");
        var result = edit.SaveBytesResult(AcceptCaches()); var saved = IWorkSourceDocument.Open(result.GetPackageBytes());
        var projected = saved.ReadPages();
        Assert.True(projected.HasEditableContent); Assert.Empty(projected.Diagnostics);
        Assert.Equal("Bravo café 😃", projected.Paragraphs[0]); Assert.Equal("Edited text", projected.Paragraphs[1]);
        Assert.Equal(body.Paragraphs.Count, projected.Body.Paragraphs.Count);
        Assert.Equal("Alpha café 😀", body.Paragraphs[0].Text); Assert.Equal("Second body", body.Paragraphs[1].Text);
        Assert.Equal(body.SourceIdentity!.RecordIdentifier, Assert.Single(result.Report.UpdatedRecords).RecordIdentifier);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "IWORK_NATIVE_VISUAL_CACHES_UNASSESSED"
            && item.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
        Assert.Equal(source.Records.Count, saved.Records.Count);
        foreach (var record in source.Records) {
            var output = saved.Records.Single(item => item.Identifier == record.Identifier && item.PayloadIndex == record.PayloadIndex);
            Assert.Equal(record.EntryPath, output.EntryPath); Assert.Equal(record.PayloadLength, output.PayloadLength);
            Assert.Equal(record.Versions, output.Versions); Assert.Equal(record.ObjectReferences, output.ObjectReferences);
            Assert.Equal(record.DataReferences, output.DataReferences);
            if (record.Identifier != body.SourceIdentity.RecordIdentifier) Assert.Equal(record.GetPayload(), output.GetPayload());
        }
        foreach (var entry in source.Entries.Where(item => !result.Report.UpdatedEntryPaths.Contains(item.Path)))
            Assert.Equal(entry.GetBytes(), saved.Entries.Single(item => item.Path == entry.Path).GetBytes());
        Assert.Equal(body.Paragraphs[0].Runs[0].Style.FontName, projected.Body.Paragraphs[0].Runs[0].Style.FontName);
        Assert.Equal(body.Paragraphs[0].Runs[0].Style.Color!.RgbHex, projected.Body.Paragraphs[0].Runs[0].Style.Color!.RgbHex);
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(result.GetPackageBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream(); converted.Value.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Contains(reopened.Paragraphs, item => item.Text == "Bravo café 😃");
        Assert.Contains(reopened.Paragraphs, item => item.Text == "Edited text");
    }

    [Fact]
    public void Invalid_or_cancelled_replacements_leave_pending_edits_intact_and_reset_restores_original_bytes() {
        var source = Source(); var body = source.ReadPages().Body; var id = body.SourceIdentity!; var edit = source.BeginEdit();
        edit.SetPagesBodyParagraphText(id, 1, "Bravo café 😃");
        edit.SetPagesBodyParagraphText(id, 2, "Edited text");
        byte[] before = edit.SaveBytes(AcceptCaches());
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(id, 2, "short"));
        // Both replacements contain eleven UTF-16 units, but the second has twelve UTF-8 bytes.
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(id, 2, "Edited téxt"));
        Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetPagesBodyParagraphText(id, 0, "Edited text"));
        Assert.Throws<ArgumentOutOfRangeException>(() => edit.SetPagesBodyParagraphText(id, body.Paragraphs.Count + 1, ""));
        Assert.Throws<ArgumentException>(() => edit.SetPagesBodyParagraphText(source.ReadPages().SourceIdentity!, 2, "Edited text"));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => edit.SetPagesBodyParagraphText(id, 2, "Second body", cancellation.Token));
        Assert.Equal(before, edit.SaveBytes(AcceptCaches()));
        using var output = new MemoryStream(); output.WriteByte(9);
        Assert.Throws<InvalidOperationException>(() => edit.Save(output)); Assert.Equal(new byte[] { 9 }, output.ToArray());
        edit.SetPagesBodyParagraphText(id, 1, body.Paragraphs[0].Text);
        Assert.Single(edit.SaveBytesResult(AcceptCaches()).Report.UpdatedRecords);
        edit.SetPagesBodyParagraphText(id, 2, body.Paragraphs[1].Text);
        var restored = edit.SaveBytesResult(); Assert.Empty(restored.Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), restored.Report.OutputPackageContentHash);
    }

    [Theory]
    [InlineData("Edited\ntext")]
    [InlineData("Edited\rtext")]
    [InlineData("Edited\u2028text")]
    [InlineData("Edited\u2029text")]
    [InlineData("Edited\ufffctext")]
    public void Paragraph_separators_inline_markers_and_invalid_Unicode_are_rejected_without_mutation(string replacement) {
        var source = Source(); var edit = source.BeginEdit();
        Assert.Throws<ArgumentException>(() => edit.SetPagesBodyParagraphText(source.ReadPages().Body.SourceIdentity!, 2, replacement));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Unpaired_surrogates_are_rejected_before_encoding() {
        var source = Source(); var edit = source.BeginEdit();
        // Construct at runtime: custom-attribute serialization replaces an unpaired UTF-16 value.
        string replacement = "Edited" + new string('\ud800', 1) + "text";
        Assert.Throws<ArgumentException>(() => edit.SetPagesBodyParagraphText(source.ReadPages().Body.SourceIdentity!, 2, replacement));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Mixed_runs_and_body_attachments_require_another_edit_profile() {
        var model = IWorkPagesDocument.Create();
        model.AddRichParagraph(new[] { new IWorkTextRunData("Alpha", bold: true), new IWorkTextRunData(" beta") });
        Reject(model, "Other text");
        model = IWorkPagesDocument.Create();
        model.CreateNumberedList().AddRichItem(new[] { new IWorkTextRunData("Alpha", "Arial", 12, bold: true), new IWorkTextRunData(" beta", "Arial", 12) });
        Reject(model, "Other text");
        model = IWorkPagesDocument.Create(); model.AddParagraph("Alpha beta"); model.AddTable(IWorkTableData.Create(1, 1).SetText(1, 1, "Cell"));
        Reject(model, "Other text");
    }

    [Fact]
    public void Nested_index_and_source_projection_limits_reject_before_pending_state_changes() {
        var original = Source(); var body = original.ReadPages().Body;
        using var inner = new MemoryStream();
        using (var index = new ZipArchive(inner, ZipArchiveMode.Create, true))
            foreach (var entry in original.Entries.Where(item => item.Path.StartsWith("Index/", StringComparison.Ordinal)))
                using (var stream = index.CreateEntry(entry.Path).Open()) { byte[] data = entry.GetBytes(); stream.Write(data); }
        using var outer = new MemoryStream();
        using (var package = new ZipArchive(outer, ZipArchiveMode.Create, true)) {
            using (var stream = package.CreateEntry("Index.zip").Open()) { stream.Write(inner.ToArray()); }
            foreach (var entry in original.Entries.Where(item => !item.Path.StartsWith("Index/", StringComparison.Ordinal)))
                using (var stream = package.CreateEntry(entry.Path).Open()) { byte[] data = entry.GetBytes(); stream.Write(data); }
        }
        var nested = IWorkSourceDocument.Open(outer.ToArray(), IWorkDocumentKind.Pages); var edit = nested.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(nested.ReadPages().Body.SourceIdentity!, 2, "Edited text"));
        Assert.Equal(nested.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
        var bounded = IWorkSourceDocument.Open(original.BeginEdit().SaveBytes(), options: new() { MaximumProjectedTextCharacters = 1 });
        edit = bounded.BeginEdit();
        Assert.Throws<InvalidDataException>(() => edit.SetPagesBodyParagraphText(body.SourceIdentity!, 2, "Edited text"));
        Assert.Equal(bounded.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    private static void Reject(IWorkPagesDocument model, string text) {
        var source = IWorkSourceDocument.Open(model.SaveBytes()); var edit = source.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(source.ReadPages().Body.SourceIdentity!, 1, text));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    private static IWorkSourceDocument Source() {
        var model = IWorkPagesDocument.Create(); model.AddParagraph("Alpha café 😀", fontSizePoints: 14, color: "123456");
        model.AddParagraph("Second body"); model.AddParagraph(""); return IWorkSourceDocument.Open(model.SaveBytes());
    }

    private static IWorkNativeSaveOptions AcceptCaches() => new() { AllowUnassessedVisualCaches = true };
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
