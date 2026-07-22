using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordSoundCollection = 0x07E4;
        private const ushort RecordSoundCollectionAtom = 0x07E5;
        private const ushort RecordSound = 0x07E6;
        private const ushort RecordSoundDataBlob = 0x07E7;

        internal static byte[] BuildSoundCollectionRecord(
            LegacyPptWriterSoundCatalog catalog) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (catalog.Sounds.Count == 0) return Array.Empty<byte>();
            var seedPayload = new byte[4];
            WriteInt32(seedPayload, 0, checked((int)catalog.Sounds.Max(sound =>
                sound.Id)));
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0, RecordSoundCollectionAtom,
                    seedPayload)
            };
            children.AddRange(catalog.Sounds.Select(BuildSoundRecord));
            return BuildContainer(RecordSoundCollection, instance: 5, children);
        }

        internal static byte[] BuildSoundRecord(LegacyPptWriterSound sound) {
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0, RecordCString,
                    Encoding.Unicode.GetBytes(sound.Name)),
                BuildRecord(version: 0, instance: 1, RecordCString,
                    Encoding.Unicode.GetBytes(sound.Extension)),
                BuildRecord(version: 0, instance: 2, RecordCString,
                    Encoding.Unicode.GetBytes(sound.Id.ToString(
                        System.Globalization.CultureInfo.InvariantCulture)))
            };
            if (sound.BuiltInId.HasValue) {
                children.Add(BuildRecord(version: 0, instance: 3, RecordCString,
                    Encoding.Unicode.GetBytes(sound.BuiltInId.Value.ToString(
                        System.Globalization.CultureInfo.InvariantCulture))));
            }
            children.Add(BuildRecord(version: 0, instance: 0,
                RecordSoundDataBlob, sound.DataBytes));
            return BuildContainer(RecordSound, instance: 0, children);
        }

        internal sealed class LegacyPptWriterSoundCatalog {
            private readonly Dictionary<string, LegacyPptWriterSound> _soundsByKey =
                new(StringComparer.Ordinal);
            private readonly List<LegacyPptWriterSound> _sounds = new();
            private readonly List<LegacyPptWriterSound> _newSounds = new();
            private uint _nextId;
            private long _totalSoundBytes;

            internal LegacyPptWriterSoundCatalog() { }

            internal LegacyPptWriterSoundCatalog(
                IEnumerable<LegacyPptSound> sourceSounds, uint? soundIdSeed) {
                if (sourceSounds == null) {
                    throw new ArgumentNullException(nameof(sourceSounds));
                }
                foreach (LegacyPptSound source in sourceSounds.OrderBy(sound =>
                             sound.Id)) {
                    if (!source.HasData || source.ContentType == null) continue;
                    string extension = NormalizeSoundExtension(source.ContentType,
                        source.Extension);
                    var sound = new LegacyPptWriterSound(source.Id, source.Name,
                        extension, source.BuiltInId, source.DataBytes,
                        isExisting: true);
                    _sounds.Add(sound);
                    _totalSoundBytes = checked(_totalSoundBytes
                        + source.DataBytes.Length);
                    string key = CreateSoundKey(sound.Name, sound.Extension,
                        sound.BuiltInId, sound.DataBytes);
                    if (!_soundsByKey.ContainsKey(key)) {
                        _soundsByKey.Add(key, sound);
                    }
                    _nextId = Math.Max(_nextId, sound.Id);
                }
                _nextId = Math.Max(_nextId, soundIdSeed.GetValueOrDefault());
            }

            internal IReadOnlyList<LegacyPptWriterSound> Sounds =>
                new ReadOnlyCollection<LegacyPptWriterSound>(_sounds);

            internal IReadOnlyList<LegacyPptWriterSound> NewSounds =>
                new ReadOnlyCollection<LegacyPptWriterSound>(_newSounds);

            internal bool TryGetOrAdd(OpenXmlPart ownerPart, OpenXmlElement soundElement,
                out LegacyPptWriterSound? sound, out string? reason) {
                sound = null;
                reason = null;
                string relationshipId;
                string name;
                bool builtIn;
                if (soundElement is P.Sound transitionSound) {
                    relationshipId = transitionSound.Embed?.Value ?? string.Empty;
                    name = transitionSound.Name?.Value ?? string.Empty;
                    builtIn = transitionSound.BuiltIn?.Value == true;
                } else if (soundElement is A.HyperlinkSound hyperlinkSound) {
                    relationshipId = hyperlinkSound.Embed?.Value ?? string.Empty;
                    name = hyperlinkSound.Name?.Value ?? string.Empty;
                    builtIn = hyperlinkSound.BuiltIn?.Value == true;
                } else {
                    reason = $"Element '{soundElement.LocalName}' is not a supported embedded sound reference.";
                    return false;
                }
                if (string.IsNullOrEmpty(relationshipId)) {
                    reason = "An embedded transition or action sound has no audio relationship.";
                    return false;
                }
                DataPartReferenceRelationship[] relationships = ownerPart
                    .DataPartReferenceRelationships.Where(candidate => string.Equals(
                        candidate.Id, relationshipId, StringComparison.Ordinal)).ToArray();
                if (relationships.Length != 1
                    || relationships[0] is not AudioReferenceRelationship
                    || relationships[0].DataPart is not MediaDataPart mediaPart) {
                    reason = $"Audio relationship '{relationshipId}' is missing or is not an embedded audio data part.";
                    return false;
                }
                string extension = mediaPart.ContentType.ToLowerInvariant() switch {
                    "audio/wav" => ".wav",
                    "audio/x-wav" => ".wav",
                    "audio/aiff" => ".aif",
                    "audio/x-aiff" => ".aif",
                    _ => string.Empty
                };
                if (extension.Length == 0) {
                    reason = $"Binary transition and action sounds require WAV or AIFF audio; '{mediaPart.ContentType}' is not representable.";
                    return false;
                }
                if (name.Length == 0) name = "Sound";
                byte[] bytes;
                if (!TryReadSoundBytes(mediaPart, out bytes, out reason)) {
                    return false;
                }
                if (bytes.Length == 0) {
                    reason = "An embedded transition or action sound has an empty audio payload.";
                    return false;
                }
                if (builtIn) {
                    LegacyPptWriterSound[] matches = _sounds.Where(candidate =>
                            candidate.BuiltInId.HasValue
                            && string.Equals(CreateSoundMediaKey(candidate.Name,
                                    candidate.Extension, candidate.DataBytes),
                                CreateSoundMediaKey(name, extension, bytes),
                                StringComparison.Ordinal))
                        .ToArray();
                    if (matches.Length == 1) {
                        sound = matches[0];
                        return true;
                    }
                    reason = matches.Length == 0
                        ? "A fresh Open XML built-in sound exposes no numeric PowerPoint 97-2003 built-in identifier; this ambiguous conversion is blocked."
                        : "The embedded built-in sound matches multiple binary built-in identifiers and cannot be mapped unambiguously.";
                    return false;
                }
                string key = CreateSoundKey(name, extension, builtInId: null, bytes);
                if (_soundsByKey.TryGetValue(key, out sound)) return true;
                if (_nextId >= int.MaxValue) {
                    reason = "The binary sound identifier range is exhausted.";
                    return false;
                }
                sound = new LegacyPptWriterSound(
                    ++_nextId, name, extension,
                    builtInId: null, bytes, isExisting: false);
                _sounds.Add(sound);
                _newSounds.Add(sound);
                _soundsByKey.Add(key, sound);
                _totalSoundBytes = checked(_totalSoundBytes + bytes.Length);
                return true;
            }

            internal bool TryGetOrAddMedia(SlidePart ownerPart,
                PowerPointMedia media, out LegacyPptWriterSound? sound,
                out string? reason) {
                if (ownerPart == null) {
                    throw new ArgumentNullException(nameof(ownerPart));
                }
                if (media == null) throw new ArgumentNullException(nameof(media));
                sound = null;
                reason = null;
                string relationshipId = media.MediaReferenceId
                    ?? string.Empty;
                DataPartReferenceRelationship[] relationships = ownerPart
                    .DataPartReferenceRelationships.Where(candidate =>
                        string.Equals(candidate.Id, relationshipId,
                            StringComparison.Ordinal)).ToArray();
                if (relationships.Length != 1
                    || relationships[0] is not AudioReferenceRelationship
                    || relationships[0].DataPart is not MediaDataPart mediaPart) {
                    reason = $"Audio relationship '{relationshipId}' is missing or is not an embedded audio data part.";
                    return false;
                }
                if (!string.Equals(mediaPart.ContentType, "audio/wav",
                        StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(mediaPart.ContentType, "audio/x-wav",
                        StringComparison.OrdinalIgnoreCase)) {
                    reason = $"Binary embedded media requires WAV audio; '{mediaPart.ContentType}' is not representable.";
                    return false;
                }
                byte[] bytes;
                if (!TryReadSoundBytes(mediaPart, out bytes, out reason)) {
                    return false;
                }
                if (bytes.Length == 0) {
                    reason = "An embedded media shape has an empty audio payload.";
                    return false;
                }
                string name = string.IsNullOrWhiteSpace(media.Name)
                    ? "Audio"
                    : media.Name!;
                string key = CreateSoundKey(name, ".wav",
                    builtInId: null, bytes);
                if (_soundsByKey.TryGetValue(key, out sound)) return true;
                if (_nextId >= int.MaxValue) {
                    reason = "The binary sound identifier range is exhausted.";
                    return false;
                }
                sound = new LegacyPptWriterSound(++_nextId, name, ".wav",
                    builtInId: null, bytes, isExisting: false);
                _sounds.Add(sound);
                _newSounds.Add(sound);
                _soundsByKey.Add(key, sound);
                _totalSoundBytes = checked(_totalSoundBytes + bytes.Length);
                return true;
            }

            private bool TryReadSoundBytes(MediaDataPart mediaPart,
                out byte[] bytes, out string? reason) {
                bytes = Array.Empty<byte>();
                reason = null;
                try {
                    using Stream input = mediaPart.GetStream(FileMode.Open,
                        FileAccess.Read);
                    bytes = OfficeStreamReader.ReadAllBytes(input,
                        MaximumSoundBytes);
                    long nextTotal = checked(_totalSoundBytes + bytes.Length);
                    if (nextTotal > MaximumTotalSoundBytes) {
                        bytes = Array.Empty<byte>();
                        reason = $"Binary PowerPoint sound payloads cannot exceed {MaximumTotalSoundBytes} aggregate bytes.";
                        return false;
                    }
                    return true;
                } catch (Exception exception) when (exception
                    is IOException or InvalidDataException
                        or NotSupportedException or OverflowException) {
                    bytes = Array.Empty<byte>();
                    reason = $"The embedded audio payload cannot be read within the {MaximumSoundBytes}-byte safety limit: {exception.Message}";
                    return false;
                }
            }

            private static string CreateSoundKey(string name, string extension,
                int? builtInId, byte[] bytes) {
                byte[] hash;
                using (SHA256 algorithm = SHA256.Create()) {
                    hash = algorithm.ComputeHash(bytes);
                }
                string normalizedExtension = extension.TrimEnd('\0')
                    .TrimStart('.').ToLowerInvariant() switch {
                        "wave" => "wav",
                        "aiff" => "aif",
                        string value => value
                    };
                return name + "\0" + normalizedExtension + "\0"
                    + (builtInId?.ToString(
                        System.Globalization.CultureInfo.InvariantCulture) ?? "-")
                    + "\0" + Convert.ToBase64String(hash);
            }

            private static string CreateSoundMediaKey(string name,
                string extension, byte[] bytes) =>
                CreateSoundKey(name, extension, builtInId: null, bytes);

            private static string NormalizeSoundExtension(string contentType,
                string? extension) {
                string value = (extension ?? string.Empty).TrimEnd('\0');
                if (value.Length == 4) return value;
                return contentType == "audio/aiff" ? ".aif" : ".wav";
            }
        }

        internal sealed class LegacyPptWriterSound {
            internal LegacyPptWriterSound(uint id, string name, string extension,
                int? builtInId, byte[] dataBytes, bool isExisting) {
                Id = id;
                Name = name;
                Extension = extension;
                BuiltInId = builtInId;
                DataBytes = dataBytes;
                IsExisting = isExisting;
            }

            internal uint Id { get; }
            internal string Name { get; }
            internal string Extension { get; }
            internal int? BuiltInId { get; }
            internal byte[] DataBytes { get; }
            internal bool IsExisting { get; }
        }
    }
}
