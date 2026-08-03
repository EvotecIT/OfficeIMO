using System;
using System.IO;
using System.IO.Packaging;
using System.Text;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioSignatureTests {
        private const string OriginRelationshipType =
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        private const string SignatureRelationshipType =
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

        [Fact]
        public void SignedVsdxBlocksRebuildUntilInvalidatedCarrierRemovalIsExplicit() {
            string sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            try {
                VisioDocument created = VisioDocument.Create(sourcePath);
                created.AddPage("Page-1").Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Signed"));
                created.Save();
                AddSignatureCarrier(sourcePath);

                VisioDocument loaded = VisioDocument.Load(sourcePath);
                VisioSignatureInfo info = loaded.InspectSignatures();
                Assert.True(info.HasSignatures);
                Assert.Equal(1, info.OriginRelationshipCount);
                Assert.Equal(1, info.OriginPartCount);
                Assert.Equal(1, info.XmlSignaturePartCount);

                VisioSignedDocumentMutationException blocked = Assert.Throws<VisioSignedDocumentMutationException>(
                    () => loaded.ToBytes());
                Assert.Same(info, blocked.SignatureInfo);

                loaded.SignatureMutationPolicy = VisioSignatureMutationPolicy.RemoveInvalidatedSignatures;
                byte[] rebuilt = loaded.ToBytes();
                using var stream = new MemoryStream(rebuilt, writable: false);
                using Package package = Package.Open(stream, FileMode.Open, FileAccess.Read);
                Assert.Empty(package.GetRelationshipsByType(OriginRelationshipType));
                Assert.DoesNotContain(package.GetParts(), part =>
                    string.Equals(
                        part.ContentType,
                        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                        StringComparison.OrdinalIgnoreCase));
            } finally {
                if (File.Exists(sourcePath)) File.Delete(sourcePath);
            }
        }

        private static void AddSignatureCarrier(string filePath) {
            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            var originUri = new Uri("/_xmlsignatures/origin.sigs", UriKind.Relative);
            PackagePart origin = package.CreatePart(
                originUri,
                "application/vnd.openxmlformats-package.digital-signature-origin");
            package.CreateRelationship(originUri, TargetMode.Internal, OriginRelationshipType, "rIdSignatureOrigin");

            var signatureUri = new Uri("/_xmlsignatures/sig1.xml", UriKind.Relative);
            PackagePart signature = package.CreatePart(
                signatureUri,
                "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml");
            origin.CreateRelationship(new Uri("sig1.xml", UriKind.Relative), TargetMode.Internal,
                SignatureRelationshipType, "rIdSignature1");
            using Stream output = signature.GetStream(FileMode.Create, FileAccess.Write);
            byte[] xml = Encoding.UTF8.GetBytes("<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />");
            output.Write(xml, 0, xml.Length);
        }
    }
}
