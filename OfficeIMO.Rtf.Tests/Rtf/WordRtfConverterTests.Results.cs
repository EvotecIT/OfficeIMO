using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Rtf_ToWord_Result_Reports_Styles_Lists_Objects_And_Shapes() {
        RtfDocument rtf = RtfDocument.Create();
        rtf.AddStyle(7, "Clinical");
        rtf.AddListDefinition(10, "Steps");
        rtf.AddListOverride(20, 10);
        rtf.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3 });
        rtf.AddShape().AddTextBoxParagraph("Shape text");

        RtfConversionResult<WordDocument> result = rtf.ToWordDocumentResult();
        using (result.Value) {
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordStylesFlattened" && diagnostic.Action == RtfConversionAction.Flattened);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordListDefinitionsFlattened" && diagnostic.Count == 2);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordObjectsOmitted" && diagnostic.Action == RtfConversionAction.Omitted);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordShapesOmitted" && diagnostic.Action == RtfConversionAction.Omitted);
            Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
        }
    }

    [Fact]
    public void Word_Rtf_Read_Result_Combines_Core_Policy_And_Bridge_Diagnostics() {
        const string rtf = @"{\rtf1{\object\objemb{\*\objdata 0102}}Visible}";

        RtfConversionResult<WordDocument> result = rtf.LoadFromRtfResult();
        using (result.Value) {
            RtfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "RTF105");
            Assert.Equal(RtfConversionAction.Blocked, diagnostic.Action);
            Assert.Equal("string", diagnostic.SourcePath);
        }
    }

    [Fact]
    public async Task Word_Rtf_Async_Read_Result_Uses_Bounded_Core_Profile() {
        var options = RtfReadOptions.CreateUntrustedProfile();
        options.MaxInputCharacters = 4;

        await Assert.ThrowsAsync<RtfReadLimitException>(() =>
            @"{\rtf1 Too large}".LoadFromRtfResultAsync(options));
    }
}
