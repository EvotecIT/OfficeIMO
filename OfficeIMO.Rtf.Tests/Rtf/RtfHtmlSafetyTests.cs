using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlSafetyTests {
    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("file:///C:/private/secret.txt")]
    [InlineData("\\\\server\\share\\secret.txt")]
    public void Default_Profile_Omits_Unsafe_Run_Hyperlinks(string target) {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddText("Visible link").SetHyperlink(new Uri(target, UriKind.RelativeOrAbsolute));

        var options = new RtfToHtmlOptions();
        string html = document.ToHtml(options);

        Assert.Equal("<p>Visible link</p>", html);
        HtmlRtfConversionDiagnostic diagnostic = Assert.Single(options.Diagnostics);
        Assert.Equal("RtfHtmlHyperlinkRejected", diagnostic.Code);
        Assert.Equal("run.Hyperlink", diagnostic.Source);
    }

    [Fact]
    public void Default_Profile_Omits_Unsafe_Field_Hyperlink_But_Keeps_Result() {
        RtfDocument document = RtfDocument.Create();
        RtfField field = document.AddParagraph().AddField("HYPERLINK \"javascript:alert(1)\"");
        field.AddText("Visible field result");

        var options = new RtfToHtmlOptions();
        string html = document.ToHtml(options);

        Assert.Equal("<p><span>Visible field result</span></p>", html);
        HtmlRtfConversionDiagnostic diagnostic = Assert.Single(options.Diagnostics);
        Assert.Equal("RtfHtmlFieldHyperlinkRejected", diagnostic.Code);
        Assert.Equal("field.Hyperlink", diagnostic.Source);
    }

    [Fact]
    public void Default_Profile_Omits_Private_Object_Payload_And_Keeps_Display_Result() {
        RtfDocument document = RtfDocument.Create();
        RtfObject rtfObject = document.AddParagraph().AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3, 4 });
        rtfObject.ClassName = "Package";
        rtfObject.Result.AddText("Attachment").SetBold();

        var options = new RtfToHtmlOptions();
        string html = document.ToHtml(options);

        Assert.Equal("<p><span class=\"rtf-object-result\"><strong>Attachment</strong></span></p>", html);
        Assert.DoesNotContain("data-officeimo-rtf-", html, StringComparison.Ordinal);
        Assert.DoesNotContain("AQIDBA==", html, StringComparison.Ordinal);
        Assert.Equal("RtfHtmlObjectMetadataOmitted", Assert.Single(options.Diagnostics).Code);
    }

    [Fact]
    public void Default_Profile_Uses_Validated_Image_Source_Callback() {
        RtfDocument document = RtfDocument.Create();
        document.AddImage(RtfImageFormat.Png, new byte[] { 1, 2, 3 });
        var options = new RtfToHtmlOptions {
            ImageSourceResolver = _ => "https://cdn.example.test/image.png"
        };

        string html = document.ToHtml(options);

        Assert.Equal("<img src=\"https://cdn.example.test/image.png\">", html);
        Assert.Empty(options.Diagnostics);
    }

    [Fact]
    public void Image_Source_Callback_Cannot_Bypass_Url_Policy() {
        RtfDocument document = RtfDocument.Create();
        document.AddImage(RtfImageFormat.Png, new byte[] { 1, 2, 3 });
        var options = new RtfToHtmlOptions {
            ImageSourceResolver = _ => "file:///C:/private/image.png"
        };

        string html = document.ToHtml(options);

        Assert.Empty(html);
        Assert.Collection(options.Diagnostics,
            diagnostic => Assert.Equal("RtfHtmlImageSourceRejected", diagnostic.Code),
            diagnostic => Assert.Equal("RtfHtmlImageEmbeddingDisabled", diagnostic.Code));
    }

    [Fact]
    public void Data_Uri_Embedding_Enforces_Per_Image_Limit() {
        RtfDocument document = RtfDocument.Create();
        document.AddImage(RtfImageFormat.Png, new byte[] { 1, 2, 3, 4 });
        var options = RtfToHtmlOptions.CreateRoundTripProfile();
        options.MaxEmbeddedImageBytes = 3;

        string html = document.ToHtml(options);

        Assert.Empty(html);
        Assert.Equal("RtfHtmlImageEmbeddingLimitExceeded", Assert.Single(options.Diagnostics).Code);
    }

    [Fact]
    public void RoundTrip_Profile_Explicitly_Enables_Private_Metadata_And_Binary_Data() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3, 4 });

        string html = document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());

        Assert.Contains("data-officeimo-rtf-object-data=\"AQIDBA==\"", html, StringComparison.Ordinal);
    }
}
