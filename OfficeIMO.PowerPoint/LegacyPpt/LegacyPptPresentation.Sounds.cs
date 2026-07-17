using System.Globalization;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private void ParseSoundCollection(LegacyPptRecord document,
            LegacyPptImportOptions options) {
            LegacyPptRecord[] collections = document.Children.Where(record =>
                record.Type == RecordSoundCollection).ToArray();
            if (collections.Length == 0) return;
            if (collections.Length > 1) {
                AddDiagnostic("PPT-SOUND-COLLECTION-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The document contains multiple sound collections; the first collection is used.",
                    collections[1].Offset);
            }
            LegacyPptRecord collection = collections[0];
            if (collection.Version != 0x0F || collection.Instance != 5) {
                AddDiagnostic("PPT-SOUND-COLLECTION-HEADER",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The sound collection has a nonstandard record header.",
                    collection.Offset);
            }

            LegacyPptRecord[] seedAtoms = collection.Children.Where(record =>
                record.Type == RecordSoundCollectionAtom).ToArray();
            if (seedAtoms.Length == 1 && seedAtoms[0].Version == 0
                && seedAtoms[0].Instance == 0 && seedAtoms[0].PayloadLength == 4) {
                int seed = seedAtoms[0].ReadInt32(0);
                if (seed > 0) SoundIdSeed = unchecked((uint)seed);
                else AddDiagnostic("PPT-SOUND-SEED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The sound collection identifier seed is negative.", seedAtoms[0].Offset);
            } else {
                AddDiagnostic("PPT-SOUND-SEED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The sound collection must contain one valid identifier-seed atom.",
                    collection.Offset);
            }

            foreach (LegacyPptRecord soundContainer in collection.Children.Where(record =>
                         record.Type == RecordSound)) {
                TryParseSound(soundContainer, options);
            }
            if (options.ReportUnsupportedContent && collection.Children.Any(child =>
                    child.Type != RecordSoundCollectionAtom
                    && child.Type != RecordSound)) {
                AddDiagnostic("PPT-SOUND-COLLECTION-UNKNOWN",
                    LegacyPptDiagnosticSeverity.Warning,
                    "The sound collection contains unknown records that remain preserve-only.",
                    collection.Offset);
            }
            _sounds.Sort((left, right) => left.Id.CompareTo(right.Id));
            uint greatestId = _sounds.Count == 0 ? 0 : _sounds.Max(sound => sound.Id);
            if (SoundIdSeed.HasValue && SoundIdSeed.Value < greatestId) {
                AddDiagnostic("PPT-SOUND-SEED-RANGE",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The sound collection seed {SoundIdSeed.Value} is lower than sound identifier {greatestId}.",
                    seedAtoms.Length == 1 ? seedAtoms[0].Offset : collection.Offset);
            }
        }

        private void TryParseSound(LegacyPptRecord container,
            LegacyPptImportOptions options) {
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-SOUND-HEADER", LegacyPptDiagnosticSeverity.Warning,
                    "A sound entry has a nonstandard record header and was skipped.",
                    container.Offset);
                return;
            }
            LegacyPptRecord[] names = FindSoundStrings(container, 0);
            LegacyPptRecord[] extensions = FindSoundStrings(container, 1);
            LegacyPptRecord[] ids = FindSoundStrings(container, 2);
            LegacyPptRecord[] builtInIds = FindSoundStrings(container, 3);
            LegacyPptRecord[] data = container.Children.Where(record =>
                record.Type == RecordSoundDataBlob).ToArray();
            if (!TryReadRequiredSoundString(names, "name", container, out string? name)
                || !TryReadRequiredSoundString(ids, "identifier", container,
                    out string? idText)
                || !int.TryParse(idText, NumberStyles.None, CultureInfo.InvariantCulture,
                    out int parsedId) || parsedId <= 0) {
                AddDiagnostic("PPT-SOUND-ID", LegacyPptDiagnosticSeverity.Warning,
                    "A sound entry has a missing or invalid positive identifier and was skipped.",
                    container.Offset);
                return;
            }
            uint id = unchecked((uint)parsedId);
            if (_soundsById.ContainsKey(id)) {
                AddDiagnostic("PPT-SOUND-ID-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Sound identifier {id} is defined more than once; the first definition is retained.",
                    container.Offset);
                return;
            }

            string? extension = null;
            if (extensions.Length > 0) {
                if (extensions.Length != 1 || extensions[0].Version != 0
                    || extensions[0].PayloadLength != 8) {
                    AddDiagnostic("PPT-SOUND-EXTENSION",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Sound {id} has duplicate or malformed extension data.",
                        container.Offset);
                } else {
                    extension = extensions[0].ReadUtf16Text().TrimEnd('\0');
                }
            }

            int? builtInId = null;
            if (builtInIds.Length > 0) {
                if (builtInIds.Length != 1 || builtInIds[0].Version != 0
                    || (builtInIds[0].PayloadLength & 1) != 0
                    || !int.TryParse(builtInIds[0].ReadUtf16Text().TrimEnd('\0'),
                        NumberStyles.None, CultureInfo.InvariantCulture,
                        out int parsedBuiltInId) || parsedBuiltInId < 100
                    || parsedBuiltInId > 125) {
                    AddDiagnostic("PPT-SOUND-BUILTIN-ID",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"Sound {id} has an invalid built-in sound identifier.",
                        container.Offset);
                } else {
                    builtInId = parsedBuiltInId;
                }
            }

            if (data.Length > 1) {
                AddDiagnostic("PPT-SOUND-DATA-DUPLICATE",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Sound {id} has multiple data blobs; the first payload is retained.",
                    data[1].Offset);
            }
            byte[] bytes = data.Length == 0 ? Array.Empty<byte>()
                : CopyPayload(data[0]);
            var sound = new LegacyPptSound(id, name!, extension, builtInId, bytes);
            _sounds.Add(sound);
            _soundsById.Add(id, sound);
            if (!sound.HasData) {
                AddDiagnostic("PPT-SOUND-DATA-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Sound {id} has no embedded audio payload and cannot be projected.",
                    container.Offset);
            } else if (sound.ContentType == null) {
                AddDiagnostic("PPT-SOUND-FORMAT",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Sound {id} is neither recognizable WAV nor AIFF audio and cannot be projected.",
                    container.Offset);
            }
            if (options.ReportUnsupportedContent && container.Children.Any(child =>
                    child.Type != RecordCString
                    && child.Type != RecordSoundDataBlob)) {
                AddDiagnostic("PPT-SOUND-UNKNOWN",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Sound {id} contains unknown records that remain preserve-only.",
                    container.Offset);
            }
        }

        private static LegacyPptRecord[] FindSoundStrings(LegacyPptRecord container,
            ushort instance) => container.Children.Where(record =>
                record.Type == RecordCString && record.Instance == instance).ToArray();

        private bool TryReadRequiredSoundString(LegacyPptRecord[] records,
            string field, LegacyPptRecord container, out string? value) {
            value = null;
            if (records.Length != 1 || records[0].Version != 0
                || (records[0].PayloadLength & 1) != 0) {
                AddDiagnostic("PPT-SOUND-STRING", LegacyPptDiagnosticSeverity.Warning,
                    $"A sound entry has duplicate or malformed {field} data.",
                    container.Offset);
                return false;
            }
            value = records[0].ReadUtf16Text().TrimEnd('\0');
            return value.Length > 0;
        }

        private static byte[] CopyPayload(LegacyPptRecord record) {
            byte[] recordBytes = record.CopyRecordBytes();
            var payload = new byte[record.PayloadLength];
            Buffer.BlockCopy(recordBytes, 8, payload, 0, payload.Length);
            return payload;
        }

        internal LegacyPptSound? FindSound(uint id) =>
            _soundsById.TryGetValue(id, out LegacyPptSound? sound) ? sound : null;

        private void ValidateSoundReferences() {
            foreach (LegacyPptSlide slide in _slides) {
                if (slide.Transition?.PlaySound == true) {
                    ValidateSoundReference(slide.Transition.SoundId,
                        $"slide {slide.SlideId} transition");
                }
                foreach (LegacyPptShape shape in slide.Shapes) {
                    ValidateShapeSoundReferences(shape, $"slide {slide.SlideId}");
                }
                if (slide.NotesPage != null) {
                    foreach (LegacyPptShape shape in slide.NotesPage.Shapes) {
                        ValidateShapeSoundReferences(shape,
                            $"notes page for slide {slide.SlideId}");
                    }
                }
            }
            foreach (LegacyPptMaster master in _masters) {
                foreach (LegacyPptShape shape in master.Shapes) {
                    ValidateShapeSoundReferences(shape, $"master {master.MasterId}");
                }
            }
            foreach (LegacyPptShape shape in NotesMaster?.Shapes
                     ?? Array.Empty<LegacyPptShape>()) {
                ValidateShapeSoundReferences(shape, "notes master");
            }
            foreach (LegacyPptShape shape in HandoutMaster?.Shapes
                     ?? Array.Empty<LegacyPptShape>()) {
                ValidateShapeSoundReferences(shape, "handout master");
            }
        }

        private void ValidateShapeSoundReferences(LegacyPptShape shape,
            string owner) {
            if (shape.Animation?.PlaysSound == true
                && !shape.Animation.HasSoundOverride) {
                ValidateSoundReference(shape.Animation.SoundIdReference,
                    $"{owner} shape {shape.ShapeId} animation");
            }
            foreach (LegacyPptInteraction interaction in shape.Interactions) {
                if (interaction.SoundIdReference != 0) {
                    ValidateSoundReference(interaction.SoundIdReference,
                        $"{owner} shape {shape.ShapeId} action");
                }
            }
            foreach (LegacyPptTextInteraction interaction in shape.TextBody.Interactions) {
                if (interaction.Interaction.SoundIdReference != 0) {
                    ValidateSoundReference(interaction.Interaction.SoundIdReference,
                        $"{owner} shape {shape.ShapeId} text action");
                }
            }
            foreach (LegacyPptShape child in shape.Children) {
                ValidateShapeSoundReferences(child, owner);
            }
        }

        private void ValidateSoundReference(uint id, string owner) {
            if (id == 0) {
                AddDiagnostic("PPT-SOUND-REFERENCE-NULL",
                    LegacyPptDiagnosticSeverity.Information,
                    $"The {owner} is marked to play a null sound reference; no sound was projected.",
                    null);
            } else if (!_soundsById.ContainsKey(id)) {
                AddDiagnostic("PPT-SOUND-REFERENCE-MISSING",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"The {owner} references missing sound identifier {id}.", null);
            }
        }
    }
}
