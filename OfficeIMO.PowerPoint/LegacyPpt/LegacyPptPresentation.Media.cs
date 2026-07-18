using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordExternalMediaAtom = 0x1004;
        private const ushort RecordExternalVideo = 0x1005;
        private const ushort RecordExternalAviMovie = 0x1006;
        private const ushort RecordExternalMciMovie = 0x1007;
        private const ushort RecordExternalMidiAudio = 0x100D;
        private const ushort RecordExternalCdAudio = 0x100E;
        private const ushort RecordExternalWavAudioEmbedded = 0x100F;
        private const ushort RecordExternalWavAudioLink = 0x1010;
        private const ushort RecordExternalCdAudioAtom = 0x1012;
        private const ushort RecordExternalWavAudioEmbeddedAtom = 0x1013;

        private readonly List<LegacyPptMedia> _media = new();
        private readonly Dictionary<uint, LegacyPptMedia> _mediaById = new();

        /// <summary>Gets typed binary audio and movie definitions.</summary>
        public IReadOnlyList<LegacyPptMedia> Media =>
            new ReadOnlyCollection<LegacyPptMedia>(_media);

        private void ParseMediaObjects(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            foreach (LegacyPptRecord list in document.Children.Where(record =>
                         record.Type == RecordExternalObjectList)) {
                foreach (LegacyPptRecord container in list.Children) {
                    if (container.Type is RecordExternalAviMovie
                        or RecordExternalMciMovie
                        or RecordExternalMidiAudio
                        or RecordExternalWavAudioLink) {
                        HasExternalMediaContent = true;
                    }
                    LegacyPptMedia? media = container.Type switch {
                        RecordExternalAviMovie => ReadVideoMedia(container,
                            LegacyPptMediaKind.AviMovie, options),
                        RecordExternalMciMovie => ReadVideoMedia(container,
                            LegacyPptMediaKind.MciMovie, options),
                        RecordExternalMidiAudio => ReadPathMedia(container,
                            LegacyPptMediaKind.MidiAudio, options),
                        RecordExternalWavAudioLink => ReadPathMedia(container,
                            LegacyPptMediaKind.LinkedWaveAudio, options),
                        RecordExternalCdAudio => ReadCdMedia(container,
                            options),
                        RecordExternalWavAudioEmbedded =>
                            ReadEmbeddedWaveMedia(container, options),
                        _ => null
                    };
                    if (media == null) continue;
                    if (_hyperlinksById.ContainsKey(media.Id)
                        || _oleObjectsById.ContainsKey(media.Id)
                        || _linkedOleObjectsById.ContainsKey(media.Id)
                        || _activeXControlsById.ContainsKey(media.Id)
                        || _mediaById.ContainsKey(media.Id)) {
                        AddDiagnostic("PPT-MEDIA-ID-DUPLICATE",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"Media identifier {media.Id} occurs more than once or collides with another external object; later definitions remain preserve-only.",
                            container.Offset);
                        continue;
                    }
                    _mediaById.Add(media.Id, media);
                    _media.Add(media);
                }
            }
        }

        private LegacyPptMedia? ReadVideoMedia(LegacyPptRecord container,
            LegacyPptMediaKind kind, LegacyPptImportOptions options) {
            if (!HasMediaContainerHeader(container)) {
                AddMalformedMediaDiagnostic(container, "movie container");
                return null;
            }
            LegacyPptRecord[] videos = container.Children.Where(record =>
                record.Type == RecordExternalVideo).ToArray();
            if (videos.Length != 1 || !HasMediaContainerHeader(videos[0])) {
                AddMalformedMediaDiagnostic(container, "video container");
                return null;
            }
            LegacyPptRecord video = videos[0];
            if (!TryReadMediaAtom(video, kind, out MediaAtomFields fields)) {
                return null;
            }
            if (!TryReadMediaPath(video, out string? path)) return null;
            ReportUnknownMediaChildren(video,
                new[] { RecordExternalMediaAtom, RecordCString }, options,
                fields.Id);
            ReportUnknownMediaChildren(container,
                new[] { RecordExternalVideo }, options, fields.Id);
            return CreateMedia(fields, kind, path);
        }

        private LegacyPptMedia? ReadPathMedia(LegacyPptRecord container,
            LegacyPptMediaKind kind, LegacyPptImportOptions options) {
            if (!HasMediaContainerHeader(container)
                || !TryReadMediaAtom(container, kind,
                    out MediaAtomFields fields)
                || !TryReadMediaPath(container, out string? path)) {
                if (!HasMediaContainerHeader(container)) {
                    AddMalformedMediaDiagnostic(container, "media container");
                }
                return null;
            }
            ReportUnknownMediaChildren(container,
                new[] { RecordExternalMediaAtom, RecordCString }, options,
                fields.Id);
            return CreateMedia(fields, kind, path);
        }

        private LegacyPptMedia? ReadCdMedia(LegacyPptRecord container,
            LegacyPptImportOptions options) {
            if (!HasMediaContainerHeader(container)
                || !TryReadMediaAtom(container,
                    LegacyPptMediaKind.CdAudio,
                    out MediaAtomFields fields)) {
                if (!HasMediaContainerHeader(container)) {
                    AddMalformedMediaDiagnostic(container, "CD audio container");
                }
                return null;
            }
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordExternalCdAudioAtom).ToArray();
            if (atoms.Length != 1 || atoms[0].Version != 0
                || atoms[0].Instance != 0
                || atoms[0].PayloadLength != 8) {
                AddMalformedMediaDiagnostic(container, "CD audio atom");
                return null;
            }
            uint start = atoms[0].ReadUInt32(0);
            uint end = atoms[0].ReadUInt32(4);
            ReportUnknownMediaChildren(container,
                new[] { RecordExternalMediaAtom,
                    RecordExternalCdAudioAtom }, options, fields.Id);
            return new LegacyPptMedia(fields.Id,
                LegacyPptMediaKind.CdAudio, fields.Loop, fields.Rewind,
                fields.Narration, path: null, soundId: null,
                durationMilliseconds: null, start, end, sound: null);
        }

        private LegacyPptMedia? ReadEmbeddedWaveMedia(
            LegacyPptRecord container, LegacyPptImportOptions options) {
            if (!HasMediaContainerHeader(container)
                || !TryReadMediaAtom(container,
                    LegacyPptMediaKind.EmbeddedWaveAudio,
                    out MediaAtomFields fields)) {
                if (!HasMediaContainerHeader(container)) {
                    AddMalformedMediaDiagnostic(container,
                        "embedded WAV container");
                }
                return null;
            }
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordExternalWavAudioEmbeddedAtom).ToArray();
            if (atoms.Length != 1 || atoms[0].Version != 1
                || atoms[0].Instance != 1
                || atoms[0].PayloadLength != 8) {
                AddMalformedMediaDiagnostic(container,
                    "embedded WAV atom");
                return null;
            }
            uint soundId = atoms[0].ReadUInt32(0);
            uint duration = atoms[0].ReadUInt32(4);
            if (soundId == 0 || duration > int.MaxValue) {
                AddMalformedMediaDiagnostic(container,
                    "embedded WAV reference");
                return null;
            }
            LegacyPptSound? sound = FindSound(soundId);
            if (sound == null) {
                AddDiagnostic("PPT-MEDIA-SOUND-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Embedded WAV media {fields.Id} references unavailable sound {soundId} and remains preserve-only.",
                    atoms[0].Offset);
            } else if (!string.Equals(sound.ContentType, "audio/wav",
                           StringComparison.OrdinalIgnoreCase)) {
                AddDiagnostic("PPT-MEDIA-SOUND-FORMAT",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Embedded WAV media {fields.Id} references sound {soundId}, whose payload is not recognizable WAV audio.",
                    atoms[0].Offset);
            }
            ReportUnknownMediaChildren(container,
                new[] { RecordExternalMediaAtom,
                    RecordExternalWavAudioEmbeddedAtom }, options,
                fields.Id);
            return new LegacyPptMedia(fields.Id,
                LegacyPptMediaKind.EmbeddedWaveAudio, fields.Loop,
                fields.Rewind, fields.Narration, path: null, soundId,
                checked((int)duration), cdStart: null, cdEnd: null,
                sound);
        }

        private bool TryReadMediaAtom(LegacyPptRecord container,
            LegacyPptMediaKind kind, out MediaAtomFields fields) {
            fields = default;
            LegacyPptRecord[] atoms = container.Children.Where(record =>
                record.Type == RecordExternalMediaAtom).ToArray();
            if (atoms.Length != 1 || atoms[0].Version != 0
                || atoms[0].Instance != 0
                || atoms[0].PayloadLength != 8) {
                AddMalformedMediaDiagnostic(container, "media atom");
                return false;
            }
            LegacyPptRecord atom = atoms[0];
            uint id = atom.ReadUInt32(0);
            ushort flags = atom.ReadUInt16(4);
            if (id == 0) {
                AddMalformedMediaDiagnostic(container, "media identifier");
                return false;
            }
            if ((flags & 0xFFF8) != 0) {
                AddDiagnostic("PPT-MEDIA-FLAGS-RESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Media {id} has nonzero reserved playback flags that remain preserved only.",
                    atom.Offset);
            }
            bool narration = (flags & 0x0004) != 0;
            if (narration && kind is LegacyPptMediaKind.AviMovie
                    or LegacyPptMediaKind.MciMovie) {
                AddDiagnostic("PPT-MEDIA-VIDEO-NARRATION",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Video media {id} sets the audio-only narration flag and remains preserve-only.",
                    atom.Offset);
            }
            fields = new MediaAtomFields(id, (flags & 0x0001) != 0,
                (flags & 0x0002) != 0, narration);
            return true;
        }

        private bool TryReadMediaPath(LegacyPptRecord container,
            out string? path) {
            path = null;
            LegacyPptRecord[] paths = container.Children.Where(record =>
                record.Type == RecordCString).ToArray();
            if (paths.Length == 0) return true;
            if (paths.Length != 1 || paths[0].Version != 0
                || paths[0].Instance != 0
                || (paths[0].PayloadLength & 1) != 0) {
                AddMalformedMediaDiagnostic(container, "media path");
                return false;
            }
            path = paths[0].ReadUtf16Text().TrimEnd('\0');
            return true;
        }

        private void ReportUnknownMediaChildren(LegacyPptRecord container,
            IReadOnlyCollection<ushort> knownTypes,
            LegacyPptImportOptions options, uint id) {
            if (!options.ReportUnsupportedContent
                || !container.Children.Any(child =>
                    !knownTypes.Contains(child.Type))) return;
            AddDiagnostic("PPT-MEDIA-UNKNOWN",
                LegacyPptDiagnosticSeverity.Warning,
                $"Media {id} contains unknown records that remain preserved only.",
                container.Offset);
        }

        private void AddMalformedMediaDiagnostic(LegacyPptRecord record,
            string field) => AddDiagnostic("PPT-MEDIA-MALFORMED",
            LegacyPptDiagnosticSeverity.Warning,
            $"A binary media object has a malformed {field} and remains preserve-only.",
            record.Offset);

        private static bool HasMediaContainerHeader(LegacyPptRecord record) =>
            record.Version == 0x0F && record.Instance == 0;

        private static LegacyPptMedia CreateMedia(MediaAtomFields fields,
            LegacyPptMediaKind kind, string? path) => new(fields.Id, kind,
            fields.Loop, fields.Rewind, fields.Narration, path,
            soundId: null, durationMilliseconds: null,
            cdStart: null, cdEnd: null, sound: null);

        private readonly struct MediaAtomFields {
            internal MediaAtomFields(uint id, bool loop, bool rewind,
                bool narration) {
                Id = id;
                Loop = loop;
                Rewind = rewind;
                Narration = narration;
            }

            internal uint Id { get; }
            internal bool Loop { get; }
            internal bool Rewind { get; }
            internal bool Narration { get; }
        }
    }
}
