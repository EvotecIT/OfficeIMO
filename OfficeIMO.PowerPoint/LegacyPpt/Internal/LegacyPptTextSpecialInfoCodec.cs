using System.Globalization;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Decodes TextSpecialInfoAtom language runs and maps culture tags to the
    /// TxLCID values used by classic binary PowerPoint.
    /// </summary>
    internal static class LegacyPptTextSpecialInfoCodec {
        internal const ushort RecordTextSpecialInfoAtom = 0x0FAA;

        private const uint SpellMask = 1U << 0;
        private const uint LanguageMask = 1U << 1;
        private const uint AlternativeLanguageMask = 1U << 2;
        private const uint Ppt10ExtensionMask = 1U << 5;
        private const uint BidiMask = 1U << 6;
        private const uint SmartTagMask = 1U << 9;
        private const uint ProjectedMask = SpellMask | LanguageMask
            | AlternativeLanguageMask;

        internal static LegacyPptTextBody Apply(LegacyPptTextBody textBody,
            LegacyPptRecord? record, int? rawCharacterCount = null) {
            if (textBody == null) throw new ArgumentNullException(
                nameof(textBody));
            if (record == null) return textBody;
            try {
                if (record.Version != 0 || record.Instance != 0
                    || record.Type != RecordTextSpecialInfoAtom) {
                    throw new InvalidDataException(
                        "The TextSpecialInfoAtom has an invalid record header.");
                }
                var cursor = new LegacyPptTextPropertyCursor(record,
                    "TextSpecialInfoAtom");
                int expected = rawCharacterCount.HasValue
                    ? checked(rawCharacterCount.Value + 1)
                    : checked(textBody.Text.Length + 1);
                int covered = 0;
                var runs = new List<LegacyPptTextLanguageRun>();
                bool hasUnprojected = rawCharacterCount.HasValue
                    && rawCharacterCount.Value != textBody.Text.Length;
                while (covered < expected) {
                    uint countValue = cursor.ReadUInt32();
                    if (countValue == 0 || countValue > int.MaxValue) {
                        throw new InvalidDataException(
                            "A TextSIRun has an invalid character count.");
                    }
                    int count = checked((int)countValue);
                    if (covered > expected - count) {
                        throw new InvalidDataException(
                            "TextSpecialInfoAtom runs exceed the corresponding text.");
                    }
                    uint masks = cursor.ReadUInt32();
                    bool runUnprojected = (masks & ~ProjectedMask) != 0;
                    bool? spellingError = null;
                    bool? needsRecheck = null;
                    if ((masks & SpellMask) != 0) {
                        ushort spelling = cursor.ReadUInt16();
                        if ((spelling & ~0x0007) != 0) {
                            throw new InvalidDataException(
                                "A TextSIException spelling flag is invalid.");
                        }
                        spellingError = (spelling & 0x0001) != 0;
                        needsRecheck = (spelling & 0x0002) != 0;
                        runUnprojected |= (spelling & 0x0004) != 0;
                    }
                    ushort? languageId = null;
                    string? language = null;
                    bool noProof = false;
                    if ((masks & LanguageMask) != 0) {
                        languageId = cursor.ReadUInt16();
                        language = DecodeLanguage(languageId.Value,
                            allowNoProof: true, out noProof,
                            out bool languageUnprojected);
                        runUnprojected |= languageUnprojected;
                    }
                    ushort? alternativeLanguageId = null;
                    string? alternativeLanguage = null;
                    if ((masks & AlternativeLanguageMask) != 0) {
                        alternativeLanguageId = cursor.ReadUInt16();
                        alternativeLanguage = DecodeLanguage(
                            alternativeLanguageId.Value,
                            allowNoProof: false, out _,
                            out bool alternativeUnprojected);
                        runUnprojected |= alternativeUnprojected;
                    }
                    if ((masks & BidiMask) != 0) {
                        short bidi = cursor.ReadInt16();
                        if (bidi is not 0 and not 1) {
                            throw new InvalidDataException(
                                "A TextSIException bidirectional flag is invalid.");
                        }
                    }
                    if ((masks & Ppt10ExtensionMask) != 0) {
                        cursor.ReadUInt32();
                    }
                    if ((masks & SmartTagMask) != 0) {
                        uint countSmartTags = cursor.ReadUInt32();
                        if (countSmartTags > int.MaxValue / 4U) {
                            throw new InvalidDataException(
                                "A TextSIException smart-tag collection is too large.");
                        }
                        cursor.Skip(checked((int)countSmartTags * 4));
                    }
                    runs.Add(new LegacyPptTextLanguageRun(covered, count,
                        languageId, language, alternativeLanguageId,
                        alternativeLanguage, noProof, spellingError,
                        needsRecheck, runUnprojected));
                    covered = checked(covered + count);
                    hasUnprojected |= runUnprojected;
                }
                if (covered != expected || !cursor.IsAtEnd) {
                    throw new InvalidDataException(
                        "TextSpecialInfoAtom does not cover the corresponding text exactly.");
                }
                hasUnprojected |= HasMixedExplicitNoLanguage(runs,
                        alternative: false)
                    || HasMixedExplicitNoLanguage(runs,
                        alternative: true);
                return textBody.WithLanguageInformation(runs,
                    hasTextSpecialInfoRecord: true, hasUnprojected,
                    isTextSpecialInfoTruncated: false);
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException
                    or CultureNotFoundException) {
                return textBody.WithLanguageInformation(
                    Array.Empty<LegacyPptTextLanguageRun>(),
                    hasTextSpecialInfoRecord: true,
                    hasUnprojectedTextSpecialInfo: true,
                    isTextSpecialInfoTruncated: true);
            }
        }

        internal static bool TryEncodeLanguage(string? language,
            out ushort languageId, out string? reason) {
            languageId = 0;
            reason = null;
            if (string.IsNullOrWhiteSpace(language)) return true;
            try {
                CultureInfo culture = CultureInfo.GetCultureInfo(
                    language!.Trim());
                int lcid = culture.LCID;
                if (!IsPersistableLanguageId(lcid)
                    || string.IsNullOrWhiteSpace(culture.Name)) {
                    reason = $"Language '{language}' has no classic PowerPoint LCID mapping.";
                    return false;
                }
                CultureInfo roundTripped = CultureInfo.GetCultureInfo(lcid);
                if (!string.Equals(roundTripped.Name, culture.Name,
                        StringComparison.OrdinalIgnoreCase)) {
                    reason = $"Language '{language}' has no classic PowerPoint LCID mapping.";
                    return false;
                }
                languageId = checked((ushort)lcid);
                return true;
            } catch (CultureNotFoundException) {
                reason = $"Language '{language}' has no classic PowerPoint LCID mapping.";
                return false;
            }
        }

        internal static bool IsPersistableLanguageId(int languageId) =>
            languageId > 0x0400 && languageId <= ushort.MaxValue
            && languageId != 0x1000
            && !IsTransientLanguageId(languageId);

        private static bool IsTransientLanguageId(int languageId) =>
            languageId >= 0x2000 && languageId <= 0x4C00
            && (languageId & 0x03FF) == 0;

        private static bool HasMixedExplicitNoLanguage(
            IReadOnlyList<LegacyPptTextLanguageRun> runs,
            bool alternative) {
            bool hasExplicitNoLanguage = false;
            bool allExplicitNoLanguage = runs.Count > 0;
            foreach (LegacyPptTextLanguageRun run in runs) {
                ushort? languageId = alternative
                    ? run.AlternativeLanguageId
                    : run.LanguageId;
                hasExplicitNoLanguage |= languageId == 0;
                allExplicitNoLanguage &= languageId == 0;
            }
            return hasExplicitNoLanguage && !allExplicitNoLanguage;
        }

        private static string? DecodeLanguage(ushort languageId,
            bool allowNoProof, out bool noProof,
            out bool hasUnprojectedInformation) {
            noProof = false;
            hasUnprojectedInformation = false;
            if (languageId == 0) {
                return null;
            }
            if (languageId == 0x0400) {
                if (allowNoProof) {
                    noProof = true;
                } else {
                    hasUnprojectedInformation = true;
                }
                return null;
            }
            if (languageId == 0x0013) {
                hasUnprojectedInformation = true;
                return null;
            }
            if (languageId == 0x1000
                || IsTransientLanguageId(languageId)) {
                hasUnprojectedInformation = true;
                return null;
            }
            if (languageId <= 0x0400) {
                throw new InvalidDataException(
                    $"TxLCID 0x{languageId:X4} is undefined.");
            }
            CultureInfo culture = CultureInfo.GetCultureInfo(languageId);
            if (string.IsNullOrWhiteSpace(culture.Name)) {
                hasUnprojectedInformation = true;
                return null;
            }
            return culture.Name;
        }
    }
}
