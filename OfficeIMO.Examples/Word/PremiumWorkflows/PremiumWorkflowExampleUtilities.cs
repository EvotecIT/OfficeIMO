using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static class PremiumWorkflowExampleUtilities {
        internal static void AddSyntheticSignatureMetadata(string filePath) {
            byte[] signatureBytes = Encoding.UTF8.GetBytes(
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\">" +
                "<SignedInfo>" +
                "<SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\" />" +
                "<Reference URI=\"/word/document.xml\">" +
                "<DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\" />" +
                "<DigestValue>T2ZmaWNlSU1P</DigestValue>" +
                "</Reference>" +
                "</SignedInfo>" +
                "<KeyInfo><X509Data><X509SubjectName>CN=OfficeIMO Example</X509SubjectName></X509Data></KeyInfo>" +
                "<Object><SignatureProperties><SignatureProperty Target=\"#OfficeIMOExampleSignature\">" +
                "<mdssi:SignatureTime xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\">" +
                "<mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format>" +
                "<mdssi:Value>2026-06-30T08:15:30Z</mdssi:Value>" +
                "</mdssi:SignatureTime>" +
                "</SignatureProperty></SignatureProperties></Object>" +
                "</Signature>");

            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, true);
            package.AddDigitalSignatureOriginPart();
            XmlSignaturePart signaturePart = package.DigitalSignatureOriginPart!.AddNewPart<XmlSignaturePart>();
            using (var stream = new MemoryStream(signatureBytes)) {
                signaturePart.FeedData(stream);
            }

            ExtendedFilePropertiesPart appPart = package.ExtendedFilePropertiesPart ?? package.AddExtendedFilePropertiesPart();
            appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
            appPart.Properties.DigitalSignature = new DigitalSignature(
                new VTBlob(Convert.ToBase64String(signatureBytes)));
            appPart.Properties.Save();
        }

        internal static void WriteSignaturePreflightReport(string path, WordSignatureValidationReport validationReport, string savePolicyMessage) {
            WordSignatureInfo signatureInfo = validationReport.SignatureInfo;
            var builder = new StringBuilder();
            builder.AppendLine("# Signature Preflight");
            builder.AppendLine();
            builder.AppendLine("- Has signatures: " + signatureInfo.HasSignatures);
            builder.AppendLine("- Has signature origin part: " + signatureInfo.HasDigitalSignatureOriginPart);
            builder.AppendLine("- Has application signature metadata: " + signatureInfo.HasApplicationSignatureMetadata);
            builder.AppendLine("- Signature parts: " + signatureInfo.SignatureParts.Count);
            builder.AppendLine("- Package structure: " + validationReport.PackageStructureStatus);
            builder.AppendLine("- XML signature structure: " + validationReport.XmlSignatureStatus);
            builder.AppendLine("- Cryptographic validation: " + validationReport.CryptographicStatus);
            builder.AppendLine("- Certificate-chain trust: " + validationReport.CertificateChainStatus);
            builder.AppendLine("- Revocation: " + validationReport.RevocationStatus);
            builder.AppendLine("- Timestamp: " + validationReport.TimestampStatus);
            builder.AppendLine("- Signed-part coverage: " + validationReport.SignedPartCoverageStatus);
            builder.AppendLine("- Save policy: " + savePolicyMessage);
            builder.AppendLine();

            foreach (WordSignaturePartInfo part in signatureInfo.SignatureParts) {
                builder.AppendLine("## " + part.Uri);
                builder.AppendLine();
                builder.AppendLine("- Signature method: " + (part.SignatureMethodAlgorithm ?? string.Empty));
                builder.AppendLine("- Digest methods: " + string.Join(", ", part.DigestMethodAlgorithms));
                foreach (WordSignatureTimestampInfo timestamp in part.Timestamps) {
                    builder.AppendLine("- Timestamp metadata: " + timestamp.Kind +
                                       " = " + (timestamp.Value ?? string.Empty) +
                                       (string.IsNullOrWhiteSpace(timestamp.Format) ? string.Empty : " (" + timestamp.Format + ")"));
                }

                builder.AppendLine("- X509 subjects: " + string.Join(", ", part.X509SubjectNames));
                foreach (WordSignatureReferenceInfo reference in part.SignedReferences) {
                    builder.AppendLine("- Signed reference: " + (reference.Uri ?? string.Empty) +
                                       " -> " + (reference.TargetPartUri ?? "not a package part") +
                                       " (" + (reference.TargetPartExists?.ToString() ?? "not checked") + ")" +
                                       "; digest value present: " + reference.HasDigestValue);
                }

                builder.AppendLine();
            }

            if (validationReport.Findings.Count > 0) {
                builder.AppendLine();
                builder.AppendLine("## Validation Findings");
                builder.AppendLine();
                foreach (string finding in validationReport.Findings) {
                    builder.AppendLine("- " + finding);
                }
            }

            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }
    }
}
