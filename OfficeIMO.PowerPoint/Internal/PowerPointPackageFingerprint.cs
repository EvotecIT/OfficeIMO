using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>Creates deterministic fingerprints over a presentation package and its relationships.</summary>
    internal static class PowerPointPackageFingerprint {
        private const int MaximumPartCount = 100000;
        private const int MaximumPartDepth = 256;
        private const int MaximumRelationshipCount = 1000000;
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
            HashSet<OpenXmlPart> parts = CollectParts(
                document, MaximumPartCount, MaximumPartDepth, MaximumRelationshipCount);
            using var content = new FingerprintWriter(MaximumContentBytes, MaximumRelationshipCount);

            foreach (OpenXmlPart part in parts.OrderBy(item => item.Uri.ToString(), StringComparer.Ordinal)) {
                if (includePart != null && !includePart(part)) continue;
                content.Append(part.Uri.ToString());
                content.Append(part.ContentType);
                try {
                    OpenXmlPartRootElement? root = part.RootElement;
                    if (root != null) {
                        if (normalizeRoot != null) {
                            OpenXmlElement normalized = root.CloneNode(true);
                            normalizeRoot(part, normalized);
                            content.AppendXml(normalized, MaximumNormalizedXmlBytes);
                        } else {
                            content.AppendXml(root, MaximumNormalizedXmlBytes);
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

        internal static HashSet<OpenXmlPart> CollectParts(PresentationDocument document,
            int maximumPartCount, int maximumPartDepth, int maximumRelationshipCount) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (maximumPartCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPartCount));
            if (maximumPartDepth < 0) throw new ArgumentOutOfRangeException(nameof(maximumPartDepth));
            if (maximumRelationshipCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumRelationshipCount));
            var parts = new HashSet<OpenXmlPart>();
            var scheduled = new HashSet<OpenXmlPart>();
            var pending = new Stack<(OpenXmlPart Part, int Depth)>();
            int relationshipCount = 0;

            void Schedule(OpenXmlPart part, int depth) {
                if (++relationshipCount > maximumRelationshipCount) {
                    throw new FingerprintLimitExceededException(
                        "The presentation relationship count exceeds the fingerprint limit.");
                }
                if (!scheduled.Add(part)) return;
                if (scheduled.Count > maximumPartCount) {
                    throw new FingerprintLimitExceededException(
                        "The presentation part count exceeds the fingerprint limit.");
                }
                if (depth > maximumPartDepth) {
                    throw new FingerprintLimitExceededException(
                        "The presentation part graph depth exceeds the fingerprint limit.");
                }
                pending.Push((part, depth));
            }

            foreach (IdPartPair pair in document.Parts) Schedule(pair.OpenXmlPart, 0);
            while (pending.Count > 0) {
                (OpenXmlPart part, int depth) = pending.Pop();
                parts.Add(part);
                foreach (IdPartPair child in part.Parts) Schedule(child.OpenXmlPart, depth + 1);
            }
            return parts;
        }

        private static void AppendPackageProperty(FingerprintWriter content, string name, string? value) {
            content.Append(name);
            content.Append(value ?? string.Empty);
        }

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

            internal void AppendXml(OpenXmlElement element, long maximumXmlBytes) {
                if (element == null) throw new ArgumentNullException(nameof(element));
                if (maximumXmlBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumXmlBytes));
                using var xmlHash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                var xmlStream = new BoundedHashStream(xmlHash, maximumXmlBytes);
                var settings = new XmlWriterSettings {
                    CloseOutput = false,
                    ConformanceLevel = ConformanceLevel.Fragment,
                    Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                    OmitXmlDeclaration = true
                };
                using (XmlWriter writer = XmlWriter.Create(xmlStream, settings)) {
                    element.WriteTo(writer);
                }

                Consume(xmlStream.BytesWritten);
                Append("xml");
                AppendLength(xmlStream.BytesWritten);
                _hash.AppendData(xmlHash.GetHashAndReset());
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
                if (count < 0L || count > _maximumBytes - _bytes) {
                    throw new FingerprintLimitExceededException(
                        "The presentation content exceeds the fingerprint byte limit.");
                }
                _bytes += count;
            }
        }

        private sealed class BoundedHashStream : Stream {
            private readonly IncrementalHash _hash;
            private readonly long _maximumBytes;

            internal BoundedHashStream(IncrementalHash hash, long maximumBytes) {
                _hash = hash ?? throw new ArgumentNullException(nameof(hash));
                _maximumBytes = maximumBytes;
            }

            internal long BytesWritten { get; private set; }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => BytesWritten;
            public override long Position {
                get => BytesWritten;
                set => throw new NotSupportedException();
            }

            public override void Flush() { }

            public override void Write(byte[] buffer, int offset, int count) {
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || offset > buffer.Length - count) {
                    throw new ArgumentOutOfRangeException();
                }
                if (count > _maximumBytes - BytesWritten) {
                    throw new FingerprintLimitExceededException(
                        "A presentation XML part exceeds the fingerprint normalization limit.");
                }
                if (count == 0) return;
                _hash.AppendData(buffer, offset, count);
                BytesWritten += count;
            }

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
        }

        private sealed class FingerprintLimitExceededException : IOException {
            internal FingerprintLimitExceededException(string message) : base(message) { }
        }
    }
}
