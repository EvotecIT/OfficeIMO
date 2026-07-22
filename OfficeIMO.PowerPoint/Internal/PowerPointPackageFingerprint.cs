using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>Creates deterministic fingerprints over a presentation package and its relationships.</summary>
    internal static class PowerPointPackageFingerprint {
        private const int MaximumPartCount = 100000;
        private const int MaximumPartDepth = 256;
        private const int MaximumRelationshipCount = 1000000;
        private const int MaximumNormalizedXmlCharacters = 64 * 1024 * 1024;
        private const long MaximumNormalizedXmlBytes = 64L * 1024 * 1024;
        private const long MaximumContentBytes = 512L * 1024 * 1024;

        internal static string Create(PresentationDocument document,
            Action<OpenXmlPart, OpenXmlElement>? normalizeRoot = null,
            Func<OpenXmlPart, bool>? includePart = null,
            Func<OpenXmlPart, IdPartPair, bool>? includeRelationship = null,
            Func<OpenXmlPart, ReferenceRelationship, bool>? includeReferenceRelationship = null,
            Func<OpenXmlPart, DataPartReferenceRelationship, bool>?
                includeDataPartReferenceRelationship = null,
            bool includePackageProperties = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            HashSet<OpenXmlPart> parts = CollectParts(document);
            using var content = new FingerprintWriter(MaximumContentBytes, MaximumRelationshipCount);

            foreach (OpenXmlPart part in parts.OrderBy(item => item.Uri.ToString(), StringComparer.Ordinal)) {
                if (includePart != null && !includePart(part)) continue;
                content.Append(part.Uri.ToString());
                content.Append(part.ContentType);
                try {
                    if (IsXmlContentType(part.ContentType)) {
                        using (Stream source = part.GetStream(FileMode.Open, FileAccess.Read)) {
                            if (source.CanSeek && source.Length - source.Position > MaximumNormalizedXmlBytes) {
                                throw new FingerprintLimitExceededException(
                                    "A presentation XML part exceeds the fingerprint normalization limit.");
                            }
                        }
                    }
                    OpenXmlPartRootElement? root = part.RootElement;
                    if (root != null) {
                        if (normalizeRoot != null) {
                            OpenXmlElement normalized = root.CloneNode(true);
                            normalizeRoot(part, normalized);
                            string xml = normalized.OuterXml;
                            if (xml.Length > MaximumNormalizedXmlCharacters) {
                                throw new FingerprintLimitExceededException(
                                    "A normalized presentation XML part exceeds the fingerprint limit.");
                            }
                            content.Append(xml);
                        } else {
                            string xml = root.OuterXml;
                            if (xml.Length > MaximumNormalizedXmlCharacters) {
                                throw new FingerprintLimitExceededException(
                                    "A presentation XML part exceeds the fingerprint limit.");
                            }
                            content.Append(xml);
                        }
                    } else {
                        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                        content.Append(stream);
                    }
                } catch (FingerprintLimitExceededException) {
                    throw;
                } catch (Exception exception) when (exception is InvalidDataException || exception is IOException) {
                    content.Append("unreadable");
                }

                foreach (IdPartPair relationship in part.Parts.OrderBy(item => item.RelationshipId,
                             StringComparer.Ordinal)) {
                    if (includeRelationship != null && !includeRelationship(part, relationship)) continue;
                    content.CountRelationship();
                    content.Append(relationship.RelationshipId);
                    content.Append(relationship.OpenXmlPart.Uri.ToString());
                }
                foreach (DataPartReferenceRelationship relationship in part.DataPartReferenceRelationships
                             .OrderBy(item => item.Id, StringComparer.Ordinal)) {
                    if (includeDataPartReferenceRelationship != null &&
                        !includeDataPartReferenceRelationship(part, relationship)) continue;
                    content.CountRelationship();
                    DataPart dataPart = relationship.DataPart;
                    content.Append(relationship.Id);
                    content.Append(relationship.RelationshipType);
                    content.Append(dataPart.Uri.ToString());
                    content.Append(dataPart.ContentType);
                    try {
                        using Stream stream = dataPart.GetStream(FileMode.Open, FileAccess.Read);
                        content.Append(stream);
                    } catch (FingerprintLimitExceededException) {
                        throw;
                    } catch (Exception exception) when (exception is InvalidDataException || exception is IOException) {
                        content.Append("unreadable");
                    }
                }
                IEnumerable<ReferenceRelationship> references = part.ExternalRelationships
                    .Cast<ReferenceRelationship>()
                    .Concat(part.HyperlinkRelationships)
                    .OrderBy(item => item.Id, StringComparer.Ordinal);
                foreach (ReferenceRelationship relationship in references) {
                    if (includeReferenceRelationship != null &&
                        !includeReferenceRelationship(part, relationship)) continue;
                    content.CountRelationship();
                    content.Append(relationship.Id);
                    content.Append(relationship.RelationshipType);
                    content.Append(relationship.Uri.OriginalString);
                }
            }

            if (includePackageProperties) {
                var properties = document.PackageProperties;
                content.Append("core");
                AppendPackageProperty(content, "Creator", properties.Creator);
                AppendPackageProperty(content, "Title", properties.Title);
                AppendPackageProperty(content, "Description", properties.Description);
                AppendPackageProperty(content, "Category", properties.Category);
                AppendPackageProperty(content, "ContentStatus", properties.ContentStatus);
                AppendPackageProperty(content, "ContentType", properties.ContentType);
                AppendPackageProperty(content, "Identifier", properties.Identifier);
                AppendPackageProperty(content, "Keywords", properties.Keywords);
                AppendPackageProperty(content, "Language", properties.Language);
                AppendPackageProperty(content, "Subject", properties.Subject);
                AppendPackageProperty(content, "Revision", properties.Revision);
                AppendPackageProperty(content, "LastModifiedBy", properties.LastModifiedBy);
                AppendPackageProperty(content, "Version", properties.Version);
                AppendPackageProperty(content, "Created", properties.Created?
                    .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture));
                AppendPackageProperty(content, "Modified", properties.Modified?
                    .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture));
                AppendPackageProperty(content, "LastPrinted", properties.LastPrinted?
                    .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture));
            }

            return content.Complete();
        }

        private static HashSet<OpenXmlPart> CollectParts(PresentationDocument document) {
            var parts = new HashSet<OpenXmlPart>();
            var pending = new Stack<(OpenXmlPart Part, int Depth)>();
            foreach (IdPartPair pair in document.Parts) pending.Push((pair.OpenXmlPart, 0));
            while (pending.Count > 0) {
                (OpenXmlPart part, int depth) = pending.Pop();
                if (!parts.Add(part)) continue;
                if (parts.Count > MaximumPartCount) {
                    throw new FingerprintLimitExceededException(
                        "The presentation part count exceeds the fingerprint limit.");
                }
                if (depth > MaximumPartDepth) {
                    throw new FingerprintLimitExceededException(
                        "The presentation part graph depth exceeds the fingerprint limit.");
                }
                foreach (IdPartPair child in part.Parts) pending.Push((child.OpenXmlPart, depth + 1));
            }
            return parts;
        }

        private static void AppendPackageProperty(FingerprintWriter content, string name, string? value) {
            content.Append(name);
            content.Append(value ?? string.Empty);
        }

        private static bool IsXmlContentType(string? contentType) =>
            !string.IsNullOrWhiteSpace(contentType) &&
            (contentType!.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) ||
             contentType.IndexOf("/xml", StringComparison.OrdinalIgnoreCase) >= 0);

        private sealed class FingerprintWriter : IDisposable {
            private readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            private readonly byte[] _buffer = new byte[81920];
            private readonly char[] _charBuffer = new char[16384];
            private readonly long _maximumBytes;
            private readonly int _maximumRelationships;
            private long _bytes;
            private int _relationships;

            internal FingerprintWriter(long maximumBytes, int maximumRelationships) {
                _maximumBytes = maximumBytes;
                _maximumRelationships = maximumRelationships;
            }

            internal void Append(string value) {
                value ??= string.Empty;
                int byteCount = Encoding.UTF8.GetByteCount(value);
                AppendLength(byteCount);
                Consume(byteCount);
                Encoder encoder = Encoding.UTF8.GetEncoder();
                int characterIndex = 0;
                while (characterIndex < value.Length) {
                    int characterCount = Math.Min(_charBuffer.Length, value.Length - characterIndex);
                    value.CopyTo(characterIndex, _charBuffer, 0, characterCount);
                    characterIndex += characterCount;
                    encoder.Convert(
                        _charBuffer, 0, characterCount,
                        _buffer, 0, _buffer.Length,
                        characterIndex == value.Length,
                        out _, out int bytesUsed, out _);
                    if (bytesUsed > 0) _hash.AppendData(_buffer, 0, bytesUsed);
                }
            }

            internal void Append(Stream stream) {
                long declaredLength = stream.CanSeek ? stream.Length - stream.Position : -1L;
                AppendLength(declaredLength);
                bool precharged = declaredLength >= 0L;
                if (precharged) Consume(declaredLength);
                int read;
                while ((read = stream.Read(_buffer, 0, _buffer.Length)) > 0) {
                    if (!precharged) Consume(read);
                    _hash.AppendData(_buffer, 0, read);
                }
            }

            internal void CountRelationship() {
                if (++_relationships > _maximumRelationships) {
                    throw new FingerprintLimitExceededException(
                        "The presentation relationship count exceeds the fingerprint limit.");
                }
            }

            internal string Complete() {
                return Convert.ToBase64String(_hash.GetHashAndReset());
            }

            public void Dispose() {
                _hash.Dispose();
            }

            private void AppendLength(long length) {
                byte[] lengthBytes = BitConverter.GetBytes(length);
                _hash.AppendData(lengthBytes);
            }

            private void Consume(long count) {
                _bytes += count;
                if (_bytes > _maximumBytes) {
                    throw new FingerprintLimitExceededException(
                        "The presentation content exceeds the fingerprint byte limit.");
                }
            }
        }

        private sealed class FingerprintLimitExceededException : IOException {
            internal FingerprintLimitExceededException(string message) : base(message) { }
        }
    }
}
