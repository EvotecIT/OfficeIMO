using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Threading;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_InheritedRunningStoriesRefreshAfterReferenceEdits(bool footer) {
        using WordDocument document = CreateInheritedRunningStoryControl(footer, "default", "next");
        using (var initial = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions())))
            Assert.Contains("MARK", initial.GetPage(2).Text);

        WordSection later = document.Sections[1];
        later.AddHeadersAndFooters();
        WordHeaderFooter local = footer ? later.Footer.Default! : later.Header.Default!;
        local.AddParagraph("LOCAL");
        using (var authored = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()))) {
            Assert.Contains("MARK", authored.GetPage(1).Text);
            Assert.Contains("LOCAL", authored.GetPage(2).Text);
            Assert.DoesNotContain("MARK", authored.GetPage(2).Text);
        }

        foreach (var reference in later._sectionProperties.ChildElements.Where(element =>
            footer ? element is W.FooterReference : element is W.HeaderReference).ToArray()) reference.Remove();
        using (var inherited = PdfPigDocument.Open(document.ToPdfBytes(EmptyRunningOptions()))) {
            Assert.Contains("MARK", inherited.GetPage(2).Text);
            Assert.DoesNotContain("LOCAL", inherited.GetPage(2).Text);
        }
    }

    [Fact]
    public async Task RunningStoryResolutionScopeRestoresNestedAndCapturedOperations() {
        using WordDocument first = WordDocument.Create();
        using WordDocument second = WordDocument.Create();
        WordHeader firstHeader = first.HeaderDefaultOrCreate;
        WordHeader secondHeader = second.HeaderDefaultOrCreate;
        firstHeader.AddParagraph("FIRST"); secondHeader.AddParagraph("SECOND");
        WordHeaderFooterResolutionSnapshot captured;
        using (WordHeaderFooterResolutionSnapshot.Begin(first)) {
            captured = WordHeaderFooterResolutionSnapshot.Current(first)!;
            Assert.NotNull(captured);
            await Task.Run(() => {
                using var inner = WordHeaderFooterResolutionSnapshot.Begin(second);
                Assert.Same(secondHeader, second.Sections[0].ResolveDefaultHeader());
                Assert.Null(WordHeaderFooterResolutionSnapshot.Current(first));
            });
            Assert.Same(captured, WordHeaderFooterResolutionSnapshot.Current(first));
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => {
                using var scope = WordHeaderFooterResolutionSnapshot.Begin(second, canceled.Token);
            });
            Assert.Same(captured, WordHeaderFooterResolutionSnapshot.Current(first));
        }
        Assert.Null(WordHeaderFooterResolutionSnapshot.Current(first));
        using (captured.Use())
            Assert.Same(firstHeader, first.Sections[0].ResolveDefaultHeader());
        Assert.Null(WordHeaderFooterResolutionSnapshot.Current(first));
    }
}
