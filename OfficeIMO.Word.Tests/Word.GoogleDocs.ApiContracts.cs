using System.Reflection;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void GoogleDocsPlanningApisUseBuildVocabulary() {
        string[] names = typeof(WordGoogleDocsExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Contains("BuildGoogleDocsPlan", names);
        Assert.Contains("BuildGoogleDocsBatch", names);
        Assert.DoesNotContain("CreateGoogleDocsTranslationPlan", names);
        Assert.DoesNotContain("CreateGoogleDocsBatch", names);
    }

    [Fact]
    public void GoogleDocsTextStylesKeepValidFormattingWhenRunColorIsUnparseable() {
        string path = Path.Combine(_directoryWithFiles, "GoogleDocsInvalidRunColor.docx");
        string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
        try {
            using var document = BuildGoogleDocsSampleDocument(path, imagePath);
            GoogleDocsBatch batch = document.BuildGoogleDocsBatch();
            GoogleDocsParagraphRun boldRun = batch.Requests
                .OfType<GoogleDocsInsertParagraphRequest>()
                .First()
                .Paragraph.Runs
                .Single(run => run.Bold);
            boldRun.ForegroundColorHex = "auto";

            GoogleDocsApiBatchUpdatePayload payload = GoogleDocsApiPayloadBuilder.BuildInitialBatchUpdatePayload(batch);
            GoogleDocsApiUpdateTextStyleRequestPayload style = Assert.Single(
                payload.Requests,
                request => request.UpdateTextStyle?.TextStyle.Bold == true).UpdateTextStyle!;

            Assert.Null(style.TextStyle.ForegroundColor);
            Assert.Contains("bold", style.Fields.Split(','));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
