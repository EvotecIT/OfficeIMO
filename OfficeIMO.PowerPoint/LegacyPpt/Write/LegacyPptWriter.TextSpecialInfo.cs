using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using static OfficeIMO.PowerPoint.LegacyPpt.Internal.LegacyPptXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal static bool TryBuildTextSpecialInfoRecord(
            IReadOnlyList<OfficeXmlNode> paragraphs, out byte[]? record,
            out string? reason) {
            ArgumentNullException.ThrowIfNull(paragraphs);
            record = null;
            reason = null;
            var runs = new List<LegacyPptWriterTextSpecialInfoRun>();
            bool hasMetadata = false;
            foreach (OfficeXmlNode paragraph in paragraphs) {
                foreach (OfficeXmlNode child in paragraph.Children()) {
                    if (child.Is(A, "r")) {
                        int count = NormalizeTextForWrite(
                            RunText(child)).Length;
                        if (count == 0) continue;
                        if (!TryReadTextSpecialInfo(child.Child(A, "rPr"),
                                out LegacyPptWriterTextSpecialInfo info,
                                out reason)) return false;
                        AddTextSpecialInfoRun(runs, count, info);
                        hasMetadata |= info.HasMetadata;
                    } else if (child.Is(A, "br") || child.Is(A, "fld")) {
                        if (!TryReadTextSpecialInfo(child.Child(A, "rPr"),
                                out LegacyPptWriterTextSpecialInfo info,
                                out reason)) return false;
                        AddTextSpecialInfoRun(runs, 1, info);
                        hasMetadata |= info.HasMetadata;
                    }
                }
                if (!TryReadTextSpecialInfo(paragraph.Child(A, "endParaRPr"),
                        out LegacyPptWriterTextSpecialInfo endInfo,
                        out reason)) return false;
                AddTextSpecialInfoRun(runs, 1, endInfo);
                hasMetadata |= endInfo.HasMetadata;
            }
            if (!hasMetadata) return true;

            using var payload = new MemoryStream();
            foreach (LegacyPptWriterTextSpecialInfoRun run in runs) {
                WriteUInt32(payload, checked((uint)run.Count));
                uint masks = 0;
                if (run.Info.HasSpelling) masks |= 1U;
                if (run.Info.HasLanguage) masks |= 1U << 1;
                if (run.Info.HasAlternativeLanguage) masks |= 1U << 2;
                WriteUInt32(payload, masks);
                if (run.Info.HasSpelling) {
                    WriteUInt16(payload, run.Info.SpellingFlags);
                }
                if (run.Info.HasLanguage) {
                    WriteUInt16(payload, run.Info.LanguageId);
                }
                if (run.Info.HasAlternativeLanguage) {
                    WriteUInt16(payload, run.Info.AlternativeLanguageId);
                }
            }
            record = BuildRecord(version: 0, instance: 0,
                LegacyPptTextSpecialInfoCodec.RecordTextSpecialInfoAtom,
                payload.ToArray());
            return true;
        }

        private static bool TryReadTextSpecialInfo(
            OfficeXmlNode properties,
            out LegacyPptWriterTextSpecialInfo info,
            out string? reason) {
            info = default;
            reason = null;
            if (properties.IsNone) return true;
            string? language = properties.Attribute("lang");
            string? alternativeLanguage = properties.Attribute("altLang");
            bool noProof = properties.Bool("noProof") == true;
            if (noProof && !string.IsNullOrWhiteSpace(language)) {
                reason = "A classic PowerPoint text run cannot combine an explicit language LCID with no-proofing in the same TextSIException.";
                return false;
            }
            ushort languageId = 0;
            bool hasLanguage = noProof
                || !string.IsNullOrWhiteSpace(language);
            if (noProof) {
                languageId = 0x0400;
            } else if (!LegacyPptTextSpecialInfoCodec.TryEncodeLanguage(
                           language, out languageId, out reason)) {
                return false;
            }
            if (!LegacyPptTextSpecialInfoCodec.TryEncodeLanguage(
                    alternativeLanguage, out ushort alternativeLanguageId,
                    out reason)) return false;
            bool? spellingError = properties.Bool("err");
            bool? dirty = properties.Bool("dirty");
            bool hasSpelling = spellingError.HasValue || dirty.HasValue;
            ushort spellingFlags = 0;
            if (spellingError == true) {
                spellingFlags |= 0x0001;
            }
            if (dirty == true) {
                spellingFlags |= 0x0002;
            }
            info = new LegacyPptWriterTextSpecialInfo(hasSpelling,
                spellingFlags, hasLanguage, languageId,
                !string.IsNullOrWhiteSpace(alternativeLanguage),
                alternativeLanguageId);
            return true;
        }

        private static void AddTextSpecialInfoRun(
            IList<LegacyPptWriterTextSpecialInfoRun> runs, int count,
            LegacyPptWriterTextSpecialInfo info) {
            if (count <= 0) return;
            if (runs.Count > 0 && runs[runs.Count - 1].Info.Matches(info)) {
                LegacyPptWriterTextSpecialInfoRun previous =
                    runs[runs.Count - 1];
                runs[runs.Count - 1] = new LegacyPptWriterTextSpecialInfoRun(
                    checked(previous.Count + count), info);
                return;
            }
            runs.Add(new LegacyPptWriterTextSpecialInfoRun(count, info));
        }

        private static string NormalizeTextForWrite(string? value) =>
            (value ?? string.Empty).Replace("\r\n", "\r")
                .Replace("\n", "\r");

        private readonly struct LegacyPptWriterTextSpecialInfo {
            internal LegacyPptWriterTextSpecialInfo(bool hasSpelling,
                ushort spellingFlags, bool hasLanguage, ushort languageId,
                bool hasAlternativeLanguage,
                ushort alternativeLanguageId) {
                HasSpelling = hasSpelling;
                SpellingFlags = spellingFlags;
                HasLanguage = hasLanguage;
                LanguageId = languageId;
                HasAlternativeLanguage = hasAlternativeLanguage;
                AlternativeLanguageId = alternativeLanguageId;
            }

            internal bool HasSpelling { get; }

            internal ushort SpellingFlags { get; }

            internal bool HasLanguage { get; }

            internal ushort LanguageId { get; }

            internal bool HasAlternativeLanguage { get; }

            internal ushort AlternativeLanguageId { get; }

            internal bool HasMetadata => HasSpelling || HasLanguage
                || HasAlternativeLanguage;

            internal bool Matches(LegacyPptWriterTextSpecialInfo other) =>
                HasSpelling == other.HasSpelling
                && SpellingFlags == other.SpellingFlags
                && HasLanguage == other.HasLanguage
                && LanguageId == other.LanguageId
                && HasAlternativeLanguage == other.HasAlternativeLanguage
                && AlternativeLanguageId == other.AlternativeLanguageId;
        }

        private readonly struct LegacyPptWriterTextSpecialInfoRun {
            internal LegacyPptWriterTextSpecialInfoRun(int count,
                LegacyPptWriterTextSpecialInfo info) {
                Count = count;
                Info = info;
            }

            internal int Count { get; }

            internal LegacyPptWriterTextSpecialInfo Info { get; }
        }
    }
}
