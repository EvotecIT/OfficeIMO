using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfMutationObjectMappingTests {
    [Theory]
    [InlineData("update")]
    [InlineData("reply")]
    [InlineData("state")]
    [InlineData("remove")]
    [InlineData("flatten")]
    public void AnnotationMutationMapsRetainedObjectsWithoutUsingNames(string operation) {
        byte[] source = CreateSource();
        var document = PdfDocument.Load(source);
        var annotations = document.Inspect().GetAnnotationsBySubtype("Text").OrderBy(annotation => annotation.PageNumber).ToArray();
        int first = annotations[0].ObjectNumber!.Value;
        int second = annotations[1].ObjectNumber!.Value;
        var result = operation switch {
            "update" => document.Annotations.Update(first, new() { Contents = "Changed contents", Title = "New author", Name = "duplicate" }),
            "reply" => document.Annotations.AddReply(first, "A reply", new() { Author = "Reviewer" }),
            "state" => document.Annotations.SetReviewState(first, PdfAnnotationReviewState.Completed),
            "remove" => document.Annotations.Remove(new() { ObjectNumber = second }),
            _ => document.Annotations.Flatten(new() { ObjectNumber = second })
        };
        Assert.NotNull(result.AnnotationObjectNumberMap);
        var output = result.ToDocument().Inspect();
        var retained = Assert.Single(output.Annotations, annotation => annotation.ObjectNumber == result.AnnotationObjectNumberMap![first]);
        Assert.Equal(operation == "update" ? "Changed contents" : "First comment", retained.Contents);
        Assert.Equal(1, retained.PageNumber);
        if (operation is "remove" or "flatten") {
            Assert.DoesNotContain(output.Annotations, annotation => annotation.Contents == "Second comment");
            Assert.False(result.AnnotationObjectNumberMap!.ContainsKey(second));
        }
        if (operation == "reply") {
            var reply = Assert.Single(output.Annotations, annotation => annotation.Contents == "A reply");
            Assert.Equal(retained.ObjectNumber, reply.Review!.InReplyToObjectNumber);
        }
    }

    [Fact]
    public void PageReorderMapsDuplicateNamedAnnotationsToTheirNewPages() {
        var document = PdfDocument.Load(CreateSource());
        var original = document.Inspect().GetAnnotationsBySubtype("Text").ToArray();
        var result = document.Pages.ReorderWithMapping(2, 1);
        var output = result.ToDocument().Inspect();
        foreach (var annotation in original) {
            var rewritten = Assert.Single(output.Annotations, item => item.ObjectNumber == result.ObjectNumberMap[annotation.ObjectNumber!.Value]);
            Assert.Equal(annotation.Contents, rewritten.Contents);
            Assert.Equal(annotation.PageNumber == 1 ? 2 : 1, rewritten.PageNumber);
        }
        Assert.Equal(document.Pages.Reorder(2, 1).Inspect().PageCount, output.PageCount);
        Assert.ThrowsAny<ArgumentException>(() => document.Pages.ReorderWithMapping(1, 1));
    }

    private static byte[] CreateSource() {
        byte[] bytes = PdfDocument.Create(compose => {
            compose.Page(page => page.Size(300, 400).Canvas(canvas => canvas.TextAnnotation("First comment", 40, 50)));
            compose.Page(page => page.Size(300, 400).Canvas(canvas => canvas.TextAnnotation("Second comment", 40, 50)));
        }).ToBytes();
        for (int page = 1; page <= 2; page++) {
            var document = PdfDocument.Load(bytes);
            int number = document.Inspect().GetAnnotationsBySubtype("Text").Single(annotation => annotation.PageNumber == page).ObjectNumber!.Value;
            bytes = document.Annotations.Update(number, new() { Name = "duplicate" }).Bytes;
        }
        return bytes;
    }
}
