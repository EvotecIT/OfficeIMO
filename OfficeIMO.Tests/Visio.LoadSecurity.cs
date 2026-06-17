using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLoadSecurityTests {
        [Fact]
        public void LoadRejectsDtdInCoreDocumentPart() {
            string filePath = CreateBasicVisioDocument();
            ReplaceZipEntry(filePath, "visio/document.xml", writer => {
                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<!DOCTYPE VisioDocument [<!ENTITY payload \"expanded\">]>");
                writer.Write("<VisioDocument xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\">");
                writer.Write("<DocumentSettings>&payload;</DocumentSettings>");
                writer.Write("</VisioDocument>");
            });

            Assert.ThrowsAny<XmlException>(() => VisioDocument.Load(filePath));
        }

        [Fact]
        public void LoadRejectsOversizedCorePagePartBeforeParsing() {
            string filePath = CreateBasicVisioDocument();
            ReplaceZipEntry(filePath, "visio/pages/page1.xml", writer => {
                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<PageContents xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\"><Shapes>");

                long emitted = 0;
                string chunk = new('x', 8192);
                long target = VisioDocument.MaxPackageXmlPartBytes + chunk.Length;
                while (emitted < target) {
                    writer.Write("<!--");
                    writer.Write(chunk);
                    writer.Write("-->");
                    emitted += chunk.Length + 7;
                }

                writer.Write("</Shapes></PageContents>");
            });

            Exception exception = Record.Exception(() => VisioDocument.Load(filePath));

            Assert.NotNull(exception);
            Assert.True(
                exception is InvalidDataException || exception is XmlException,
                "Expected the oversized part to be rejected by either the package byte cap or the XML character cap.");
            Assert.Contains(
                exception is InvalidDataException ? VisioDocument.MaxPackageXmlPartBytes.ToString() : "MaxCharactersInDocument",
                exception.Message);
        }

        [Fact]
        public void ListMastersRejectsDtdInMastersPart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            using (ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Create)) {
                ZipArchiveEntry entry = archive.CreateEntry("visio/masters/masters.xml", CompressionLevel.Optimal);
                using Stream stream = entry.Open();
                using StreamWriter writer = new(stream, new UTF8Encoding(false));
                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<!DOCTYPE Masters [<!ENTITY payload \"expanded\">]>");
                writer.Write("<Masters xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\">");
                writer.Write("<Master ID=\"1\" NameU=\"&payload;\" />");
                writer.Write("</Masters>");
            }

            Assert.ThrowsAny<XmlException>(() => VisioAssets.ListMasters(filePath));
        }

        [Fact]
        public void StencilCatalogManifestRejectsDtd() {
            using MemoryStream stream = new();
            using (StreamWriter writer = new(stream, new UTF8Encoding(false), 1024, leaveOpen: true)) {
                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<!DOCTYPE StencilCatalog [<!ENTITY payload \"expanded\">]>");
                writer.Write("<StencilCatalog xmlns=\"urn:officeimo:visio:stencils\" Version=\"1\" Name=\"&payload;\" />");
            }

            stream.Position = 0;

            Assert.ThrowsAny<XmlException>(() => VisioStencilCatalogManifest.Load(stream));
        }

        private static string CreateBasicVisioDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Secure Load", 8.5, 11);
            page.Shapes.Add(new VisioShape("shape-1", 1, 1, 1, 1, "Shape"));
            document.Save();
            return filePath;
        }

        private static void ReplaceZipEntry(string filePath, string entryName, Action<StreamWriter> write) {
            using ZipArchive archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            entry.Delete();

            ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using Stream stream = replacement.Open();
            using StreamWriter writer = new(stream, new UTF8Encoding(false));
            write(writer);
        }
    }
}
