#nullable enable
using DocumentFormat.OpenXml.Packaging;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    internal static partial class OfficePackageSignatureInspector {
        private static IEnumerable<string> ReadEmbeddedCertificateSubjects(
            IReadOnlyList<XElement> certificateElements,
            string signaturePartUri,
            long maxCertificateBytes,
            OfficePackageCertificateByteBudget certificateByteBudget,
            List<string> unsupportedDetails) {
            foreach (XElement element in certificateElements) {
                string certificateText = element.Value.Trim();
                if (certificateText.Length == 0) continue;

                byte[] rawCertificate;
                try {
                    if (certificateText.Length > GetMaxBase64EncodedCharacters(maxCertificateBytes)) {
                        throw new InvalidDataException("The embedded X509Certificate exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    rawCertificate = Convert.FromBase64String(certificateText);
                    if (rawCertificate.LongLength > maxCertificateBytes) {
                        throw new InvalidDataException("The embedded X509Certificate exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    certificateByteBudget.Reserve(rawCertificate.LongLength);
                } catch (FormatException exception) {
                    unsupportedDetails.Add("Unable to parse X509Certificate in XML signature part " + signaturePartUri + ": " + exception.Message);
                    continue;
                }

                string? subject = ReadCertificateSubject(rawCertificate, "embedded X509Certificate in XML signature part " + signaturePartUri, unsupportedDetails);
                if (!string.IsNullOrWhiteSpace(subject)) yield return subject!;
            }
        }

        private static IReadOnlyList<XElement> GetSignerX509DataElements(XDocument xml, XNamespace ds) {
            XElement? signature = xml.Root;
            if (signature == null) return Array.Empty<XElement>();
            return signature.Elements(ds + "KeyInfo")
                .SelectMany(keyInfo => keyInfo.Elements(ds + "X509Data"))
                .ToArray();
        }

        private static IEnumerable<string> ReadRelatedCertificateSubjects(
            XmlSignaturePart signaturePart,
            long maxCertificateBytes,
            OfficePackageCertificateByteBudget certificateByteBudget,
            List<string> unsupportedDetails) {
            foreach (IdPartPair relationship in signaturePart.Parts) {
                OpenXmlPart relatedPart = relationship.OpenXmlPart;
                if (!IsSignatureCertificatePart(relatedPart)) continue;

                byte[] rawCertificate;
                try {
                    using Stream stream = relatedPart.GetStream(FileMode.Open, FileAccess.Read);
                    if (stream.CanSeek && stream.Length > maxCertificateBytes) {
                        throw new InvalidDataException("The signature certificate part exceeds the " + maxCertificateBytes + " byte limit.");
                    }
                    using var memoryStream = new MemoryStream();
                    CopyBounded(stream, memoryStream, maxCertificateBytes);
                    rawCertificate = memoryStream.ToArray();
                    certificateByteBudget.Reserve(rawCertificate.LongLength);
                } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidOperationException) {
                    unsupportedDetails.Add("Unable to read signature certificate part " + relatedPart.Uri + ": " + exception.Message);
                    continue;
                }

                string? subject = ReadCertificateSubject(rawCertificate, "signature certificate part " + relatedPart.Uri, unsupportedDetails);
                if (!string.IsNullOrWhiteSpace(subject)) yield return subject!;
            }
        }

        private static bool IsSignatureCertificatePart(OpenXmlPart part) =>
            part.RelationshipType.EndsWith("/digital-signature/certificate", StringComparison.OrdinalIgnoreCase) ||
            part.Uri.ToString().EndsWith(".cer", StringComparison.OrdinalIgnoreCase);

        private static string? ReadCertificateSubject(byte[] rawCertificate, string source, List<string> unsupportedDetails) {
            try {
                using X509Certificate2 certificate = LoadCertificate(rawCertificate);
                string subjectName = certificate.SubjectName.Name ?? certificate.Subject;
                return string.IsNullOrWhiteSpace(subjectName) ? null : subjectName.Trim();
            } catch (CryptographicException exception) {
                unsupportedDetails.Add("Unable to parse X509 certificate from " + source + ": " + exception.Message);
                return null;
            }
        }

        private static X509Certificate2 LoadCertificate(byte[] rawCertificate) {
#if NET9_0_OR_GREATER
            return X509CertificateLoader.LoadCertificate(rawCertificate);
#else
            return new X509Certificate2(rawCertificate);
#endif
        }

        private static long GetMaxBase64EncodedCharacters(long maxDecodedBytes) =>
            maxDecodedBytes > (long.MaxValue / 4L) * 3L
                ? long.MaxValue
                : ((maxDecodedBytes + 2L) / 3L) * 4L;

        private static void CopyBounded(Stream source, Stream destination, long maxBytes) {
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
                total += read;
                if (total > maxBytes) {
                    throw new InvalidDataException("The signature certificate part exceeds the " + maxBytes + " byte limit.");
                }
                destination.Write(buffer, 0, read);
            }
        }
    }
}
