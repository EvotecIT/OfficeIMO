using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

/// <summary>
/// Ensures documents that combine complex building blocks validate against the Open XML schema.
/// </summary>
public class OpenXmlValidationTests {
    [Fact]
    public void CoverPageTableOfContentAndPageNumberAreSchemaValid() {
        using var document = WordDocument.Create();

        document.AddCoverPage(CoverPageTemplate.Austin);
        document.AddTableOfContent(TableOfContentStyle.Template2);
        document.AddPageBreak();
        document.AddParagraph("Section");
        document.AddHeadersAndFooters();
        document.Footer?.Default?.AddPageNumber(WordPageNumberStyle.Roman);

        document.Save();

        using var cloneStream = new MemoryStream();
        using var clone = document._wordprocessingDocument.Clone(cloneStream, true);
        var settingsPart = clone.MainDocumentPart?.DocumentSettingsPart;
        if (settingsPart != null) {
            clone.MainDocumentPart!.DeletePart(settingsPart);
        }

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(clone);
        Assert.Empty(errors);
    }
}
