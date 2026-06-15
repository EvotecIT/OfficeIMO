using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlDocumentReferenceTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_File_Table_And_Xml_Namespaces() {
        RtfDocument document = RtfDocument.Create();
        RtfFileReference local = document.AddFileReference(@"C:\Private\Resume\Edu\File2.docx", RtfFileSource.Ntfs);
        local.RelativePathStart = 18;
        RtfFileReference network = document.AddFileReference(@"\\Server\Share\Linked.docx", RtfFileSource.Ntfs | RtfFileSource.Network);
        network.OperatingSystemNumber = 42;
        document.AddXmlNamespace(2, "urn:contoso:custom");
        document.AddXmlNamespace(1, "http://schemas.example.test/word");
        document.AddParagraph("Body");

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<meta name=\"officeimo-rtf-file-references\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-xml-namespaces\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadRtfFromHtml();

        Assert.Collection(roundTrip.FileReferences,
            file => {
                Assert.Equal(0, file.Id);
                Assert.Equal(@"C:\Private\Resume\Edu\File2.docx", file.Path);
                Assert.Equal(18, file.RelativePathStart);
                Assert.Null(file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Ntfs, file.Sources);
            },
            file => {
                Assert.Equal(1, file.Id);
                Assert.Equal(@"\\Server\Share\Linked.docx", file.Path);
                Assert.Null(file.RelativePathStart);
                Assert.Equal(42, file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Ntfs | RtfFileSource.Network, file.Sources);
            });

        Assert.Collection(roundTrip.XmlNamespaces,
            xmlNamespace => {
                Assert.Equal(2, xmlNamespace.Id);
                Assert.Equal("urn:contoso:custom", xmlNamespace.Uri);
            },
            xmlNamespace => {
                Assert.Equal(1, xmlNamespace.Id);
                Assert.Equal("http://schemas.example.test/word", xmlNamespace.Uri);
            });

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\*\filetbl{\file\fid0\frelative18\fvalidntfs C:\\Private\\Resume\\Edu\\File2.docx}{\file\fid1\fosnum42\fvalidntfs\fnetwork \\\\Server\\Share\\Linked.docx}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\xmlnstbl{\xmlns1 http://schemas.example.test/word;}{\xmlns2 urn:contoso:custom;}}", rtf, StringComparison.Ordinal);
    }
}
