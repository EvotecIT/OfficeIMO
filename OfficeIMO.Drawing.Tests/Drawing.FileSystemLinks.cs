using System.Text.RegularExpressions;
using OfficeIMO.Internal;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeFileSystemLinksTests {
    [Fact]
    public void OrdinaryFilesAndFoldersAreNotLinks() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.FileSystemLinks", Guid.NewGuid().ToString("N"));
        try {
            string folder = Directory.CreateDirectory(Path.Combine(root, "Fonts")).FullName;
            string file = Path.Combine(folder, "Font.ttf");
            File.WriteAllBytes(file, new byte[] { 0, 1, 0, 0 });

            Assert.False(OfficeFileSystemLinks.IsLinkOrReparsePoint(new FileInfo(file)));
            Assert.False(OfficeFileSystemLinks.IsLinkOrReparsePoint(new DirectoryInfo(folder)));
            Assert.False(OfficeFileSystemLinks.IsLinkOrReparsePoint(file));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void FileAndFolderLinksAreRecognized() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.FileSystemLinks", Guid.NewGuid().ToString("N"));
        try {
            string folder = Directory.CreateDirectory(Path.Combine(root, "Fonts")).FullName;
            string file = Path.Combine(folder, "Font.ttf");
            File.WriteAllBytes(file, new byte[] { 0, 1, 0, 0 });
            string fileLink = Path.Combine(root, "Linked.ttf");
            string folderLink = Path.Combine(root, "LinkedFonts");
            try {
                File.CreateSymbolicLink(fileLink, file);
                Directory.CreateSymbolicLink(folderLink, folder);
            } catch (Exception exception) when (exception is UnauthorizedAccessException or IOException) {
                // Windows without Developer Mode can't create symbolic links; the attribute check covers it there.
                return;
            }

            Assert.True(OfficeFileSystemLinks.IsLinkOrReparsePoint(new FileInfo(fileLink)));
            Assert.True(OfficeFileSystemLinks.IsLinkOrReparsePoint(new DirectoryInfo(folderLink)));
            Assert.True(OfficeFileSystemLinks.IsLinkOrReparsePoint(fileLink));
            string folderLinkWithSeparator = folderLink + Path.DirectorySeparatorChar;
            Assert.True(OfficeFileSystemLinks.IsLinkOrReparsePoint(new DirectoryInfo(folderLinkWithSeparator)));
            Assert.True(OfficeFileSystemLinks.IsLinkOrReparsePoint(folderLinkWithSeparator));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }
#endif

    /// <summary>
    /// System font discovery reads entries the app doesn't own. On iOS, reading their attributes makes .NET check group
    /// membership, which throws OverflowException, and Word and Excel files then failed to convert. The discovery code
    /// checks links through <see cref="OfficeFileSystemLinks"/> and sizes files from the opened stream instead.
    /// </summary>
    [Theory]
    [InlineData("OfficeIMO.Pdf", "Model", "PdfEmbeddedFontFamily.System.cs")]
    [InlineData("OfficeIMO.Pdf", "Model", "PdfEmbeddedFontFamily.System.MetadataProbe.cs")]
    [InlineData("OfficeIMO.Core", "OfficeCloudFontFiles.cs")]
    [InlineData("OfficeIMO.Core", "OfficeCloudFontFiles.Catalog.cs")]
    [InlineData("OfficeIMO.Pdf", "Model", "PdfFontInput.cs")]
    public void SystemFontDiscoveryDoesNotReadEntryAttributes(params string[] relativePath) {
        string source = File.ReadAllText(Path.Combine(new[] { RepositoryTestPaths.Find() }.Concat(relativePath).ToArray()));

        Assert.DoesNotMatch(new Regex(@"\.Attributes\b(?!\s*\()"), source);
        Assert.DoesNotContain("GetAttributes(", source);
        // FileInfo.Length reads the attributes too; sizes come from the opened stream.
        Assert.DoesNotContain("FileInfo(", source);
    }
}
