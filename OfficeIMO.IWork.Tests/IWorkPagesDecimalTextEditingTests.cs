using System.Text;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesDecimalTextEditingTests {
    [Theory]
    [InlineData("pages-third-level-created.pages")]
    [InlineData("pages-third-level-saved-v15.4.pages")]
    [InlineData("pages-third-level-edited-v15.4.pages")]
    public void Fixed_length_decimal_item_edits_change_only_requested_bytes_and_preserve_saved_Word_numbering(string fixture) {
        var source = IWorkSourceDocument.Open(Fixture(fixture));
        var before = source.ReadPages().Body;
        var replacements = new Dictionary<string, string> {
            ["Parent seven"] = "Edited seven", ["Child one"] = "Child two",
            ["Grandchild one"] = "Grandchild six", ["Resumed grandchild six"] = "Updated grandchild six"
        };
        var edit = source.BeginEdit();
        for (int index = 0; index < before.Paragraphs.Count; index++)
            if (replacements.TryGetValue(before.Paragraphs[index].Text, out string? text))
                edit.SetPagesBodyParagraphText(before.SourceIdentity!, index + 1, text);
        var result = edit.SaveBytesResult(AcceptCaches());
        var saved = IWorkSourceDocument.Open(result.GetPackageBytes());
        var after = saved.ReadPages().Body;
        Assert.True(after.IsComplete); Assert.True(after.IsFormattingComplete);
        Assert.Equal(before.Paragraphs.Select(p => replacements.TryGetValue(p.Text, out string? text) ? text : p.Text),
            after.Paragraphs.Select(p => p.Text));
        Assert.Equal(before.Paragraphs.Select(p => p.ListNumber), after.Paragraphs.Select(p => p.ListNumber));
        Assert.Equal(before.Paragraphs.Select(p => p.ListLevel), after.Paragraphs.Select(p => p.ListLevel));
        Assert.Equal(before.SourceIdentity!.RecordIdentifier, Assert.Single(result.Report.UpdatedRecords).RecordIdentifier);
        Assert.Equal(source.Records.Count, saved.Records.Count);
        foreach (var original in source.Records) {
            var output = saved.Records.Single(r => r.Identifier == original.Identifier && r.PayloadIndex == original.PayloadIndex);
            Assert.Equal(original.EntryPath, output.EntryPath); Assert.Equal(original.Versions, output.Versions);
            Assert.Equal(original.ObjectReferences, output.ObjectReferences); Assert.Equal(original.DataReferences, output.DataReferences);
            byte[] expected = original.GetPayload();
            if (original.Identifier == before.SourceIdentity.RecordIdentifier && original.PayloadIndex == 0)
                foreach (var replacement in replacements) {
                    byte[] oldText = Encoding.UTF8.GetBytes(replacement.Key);
                    int at = expected.AsSpan().IndexOf(oldText);
                    Assert.True(at >= 0); Assert.Equal(-1, expected.AsSpan(at + oldText.Length).IndexOf(oldText));
                    Encoding.UTF8.GetBytes(replacement.Value).CopyTo(expected, at);
                }
            Assert.Equal(expected, output.GetPayload());
        }
        foreach (var entry in source.Entries.Where(e => !result.Report.UpdatedEntryPaths.Contains(e.Path)))
            Assert.Equal(entry.GetBytes(), saved.Entries.Single(e => e.Path == entry.Path).GetBytes());
        for (int index = 0; index < before.Paragraphs.Count; index++) {
            Assert.Equal(before.Paragraphs[index].ListLayout?.MarkerIndentPoints, after.Paragraphs[index].ListLayout?.MarkerIndentPoints);
            Assert.Equal(before.Paragraphs[index].ListLayout?.TextIndentEm, after.Paragraphs[index].ListLayout?.TextIndentEm);
            Assert.Equal(before.Paragraphs[index].Style.TextStyle.FontName, after.Paragraphs[index].Style.TextStyle.FontName);
        }
        CheckWord(saved);
        Assert.Contains(before.Paragraphs, p => p.Text == "Parent seven");
        Assert.DoesNotContain(before.Paragraphs, p => p.Text == "Edited seven");
    }

    [Fact]
    public void Unicode_decimal_edits_reject_offset_changes_and_cancellation_and_restore_original_native_bytes() {
        var model = IWorkPagesDocument.Create(); var list = model.CreateNumberedList();
        list.AddItem("Parent", restartAt: 7); list.AddItem("Child", 1); list.AddItem("Alpha café 😀", 2);
        var source = IWorkSourceDocument.Open(model.SaveBytes()); var body = source.ReadPages().Body;
        var edit = source.BeginEdit(); var id = body.SourceIdentity!;
        edit.SetPagesBodyParagraphText(id, 3, "Bravo café 😃");
        byte[] pending = edit.SaveBytes(AcceptCaches());
        Assert.Equal("Bravo café 😃", IWorkSourceDocument.Open(pending).ReadPages().Paragraphs[2]);
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(id, 3, "short"));
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(id, 3, "Bravo cafe 😃"));
        Assert.Throws<ArgumentException>(() => edit.SetPagesBodyParagraphText(id, 3, "Bravo\ncafé 😃"));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => edit.SetPagesBodyParagraphText(id, 3, "Alpha café 😀", cancellation.Token));
        Assert.Equal(pending, edit.SaveBytes(AcceptCaches()));
        Assert.Throws<InvalidOperationException>(() => edit.SaveBytes());
        edit.SetPagesBodyParagraphText(id, 3, "Alpha café 😀");
        var restored = edit.SaveBytesResult(); Assert.Empty(restored.Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), restored.Report.OutputPackageContentHash);
    }

    internal static void CheckWord(IWorkSourceDocument source) {
        var paragraphs = source.ReadPages().Body.Paragraphs;
        using var converted = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream(); converted.Value.Save(bytes);
        using (var saved = WordDocument.Load(new MemoryStream(bytes.ToArray()))) {
            Assert.Empty(saved.ValidateDocument());
            Assert.Equal(paragraphs.Where(p => p.ListNumber.HasValue).Select(p => p.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(saved).Values.Select(p => p.Index));
        }
        using var xml = WordprocessingDocument.Open(new MemoryStream(bytes.ToArray()), false);
        var output = xml.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Assert.Equal(paragraphs.Select(p => p.Text), output.Select(p => p.InnerText));
        for (int index = 0; index < paragraphs.Count; index++)
            if (paragraphs[index].ListNumber.HasValue) {
                var properties = output[index].ParagraphProperties!;
                Assert.Equal(paragraphs[index].ListLevel, properties.NumberingProperties!.NumberingLevelReference!.Val!.Value);
                Assert.Equal(((paragraphs[index].ListLevel + 1) * 360).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    properties.Indentation!.Left!.Value);
                Assert.Equal("360", properties.Indentation.Hanging!.Value);
            }
    }

    private static IWorkNativeSaveOptions AcceptCaches() => new() { AllowUnassessedVisualCaches = true };
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Literal_bullet_body_paragraphs_still_require_a_separate_native_edit_profile() {
        using var package = CreatePagesPackageWithListLabel("•");
        var source = IWorkSourceDocument.Open(package); var body = source.ReadPages().Body;
        Assert.True(body.IsComplete);
        Assert.Equal(IWorkListMarkerKind.Text, Assert.Single(body.Paragraphs).ListMarkerKind);
        var edit = source.BeginEdit();
        Assert.Throws<NotSupportedException>(() => edit.SetPagesBodyParagraphText(body.SourceIdentity!, 1, "Text"));
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }
}
