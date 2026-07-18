using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordStyleTextPropAtomForPreservation = 0x0FA1;
        private const ushort RecordTextRulerAtomForPreservation = 0x0FA6;
        private const ushort RecordTextSpecialInfoAtomForPreservation = 0x0FAA;

        private static bool TryCreateTextFontCatalog(
            LegacyPptPackage package,
            out LegacyPptWriter.LegacyPptWriterFontCatalog catalog) {
            if (!package.PersistObjects.TryGetValue(
                    package.DocumentPersistId,
                    out LegacyPptPersistObject? persistObject)) {
                catalog = LegacyPptWriter.CreateFontCatalogForWrite();
                return false;
            }
            try {
                LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                    persistObject.RecordBytes, 0,
                    new LegacyPptImportOptions());
                catalog = new LegacyPptWriter.LegacyPptWriterFontCatalog(
                    document);
                return true;
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentException) {
                catalog = LegacyPptWriter.CreateFontCatalogForWrite();
                return false;
            }
        }

        private static bool TryRewriteTextBoxFormatting(
            LegacyPptRecord textbox,
            LegacyPptWriter.LegacyPptWriterTextBoxContent content,
            bool rewriteInteractions,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterTextInteraction>
                interactions, out byte[] bytes) {
            if (textbox.Version != 0x0F
                || textbox.Children.Count(child => child.Type
                    is RecordTextChars or RecordTextBytes) != 1
                || textbox.Children.Count(child => child.Type
                    == RecordStyleTextPropAtomForPreservation) > 1
                || textbox.Children.Count(child => child.Type
                    == RecordTextRulerAtomForPreservation) > 1
                || textbox.Children.Count(child => child.Type
                    == RecordTextSpecialInfoAtomForPreservation) > 1) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            if (!TryBuildPreservedTextSpecialInfoRecord(textbox,
                    content.SpecialInfoRecord,
                    checked(content.Text.Length + 1),
                    out byte[]? specialInfoRecord)) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            byte[] textRecord = BuildRecord(version: 0, instance: 0,
                RecordTextChars, System.Text.Encoding.Unicode.GetBytes(
                    content.Text));
            var children = new List<byte[]>(textbox.Children.Count
                + interactions.Count * 2 + 2);
            bool wroteText = false;
            for (int index = 0; index < textbox.Children.Count; index++) {
                LegacyPptRecord child = textbox.Children[index];
                if (child.Type is RecordTextChars or RecordTextBytes) {
                    if (wroteText) {
                        bytes = textbox.CopyRecordBytes();
                        return false;
                    }
                    children.Add(textRecord);
                    if (content.StyleRecord != null) {
                        children.Add(content.StyleRecord);
                    }
                    children.AddRange(content.FieldRecords);
                    if (specialInfoRecord != null) {
                        children.Add(specialInfoRecord);
                    }
                    wroteText = true;
                    continue;
                }
                if (child.Type == RecordStyleTextPropAtomForPreservation
                    || child.Type == RecordTextRulerAtomForPreservation
                    || child.Type
                        == RecordTextSpecialInfoAtomForPreservation
                    || LegacyPptWriter.IsTextMetaCharacterRecord(
                        child.Type)) {
                    continue;
                }
                if (rewriteInteractions
                    && child.Type == RecordInteractiveInfo) {
                    if (!IsRewritableInteractiveInfo(child)
                        || index + 1 >= textbox.Children.Count
                        || textbox.Children[index + 1].Type
                            != RecordTextInteractiveInfoAtom
                        || textbox.Children[index + 1].Version != 0
                        || textbox.Children[index + 1].Instance
                            != child.Instance
                        || textbox.Children[index + 1].PayloadLength != 8) {
                        bytes = textbox.CopyRecordBytes();
                        return false;
                    }
                    index++;
                    continue;
                }
                if (rewriteInteractions
                    && child.Type == RecordTextInteractiveInfoAtom) {
                    bytes = textbox.CopyRecordBytes();
                    return false;
                }
                children.Add(child.CopyRecordBytes());
            }
            if (!wroteText) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            if (rewriteInteractions) {
                foreach (LegacyPptWriter.LegacyPptWriterTextInteraction
                         interaction in interactions) {
                    children.Add(LegacyPptWriter.BuildInteractiveInfoRecord(
                        interaction.Interaction));
                    children.Add(LegacyPptWriter
                        .BuildTextInteractiveInfoRecord(interaction));
                }
            }
            if (content.RulerRecord != null) {
                children.Add(content.RulerRecord);
            }
            bytes = BuildRecord(textbox.Version, textbox.Instance,
                textbox.Type, Concat(children));
            return true;
        }

        private static bool TryBuildPreservedTextSpecialInfoRecord(
            LegacyPptRecord textbox, byte[]? replacement,
            int replacementCharacterCount, out byte[]? record) {
            record = replacement;
            LegacyPptRecord? source = textbox.Children.SingleOrDefault(
                child => child.Type
                    == RecordTextSpecialInfoAtomForPreservation);
            if (source == null) return true;
            LegacyPptRecord sourceText = textbox.Children.Single(child =>
                child.Type is RecordTextChars or RecordTextBytes);
            int sourceCharacterCount = sourceText.Type == RecordTextChars
                ? sourceText.PayloadLength / 2
                : sourceText.PayloadLength;
            try {
                if (!TryReadTextSpecialInfoDefaults(source,
                        checked(sourceCharacterCount + 1),
                        out bool forcePrimaryNoLanguage,
                        out bool forceAlternativeNoLanguage)) {
                    return false;
                }
                if (!forcePrimaryNoLanguage
                    && !forceAlternativeNoLanguage) {
                    return true;
                }
                if (replacement == null) {
                    using var payload = new MemoryStream();
                    WriteUInt32ToStream(payload,
                        checked((uint)replacementCharacterCount));
                    uint masks = forcePrimaryNoLanguage ? 1U << 1 : 0;
                    if (forceAlternativeNoLanguage) masks |= 1U << 2;
                    WriteUInt32ToStream(payload, masks);
                    if (forcePrimaryNoLanguage) {
                        WriteUInt16ToStream(payload, 0);
                    }
                    if (forceAlternativeNoLanguage) {
                        WriteUInt16ToStream(payload, 0);
                    }
                    record = BuildRecord(version: 0, instance: 0,
                        RecordTextSpecialInfoAtomForPreservation,
                        payload.ToArray());
                    return true;
                }
                LegacyPptRecord generated = LegacyPptRecordReader.ReadSingle(
                    replacement, 0, new LegacyPptImportOptions());
                if (generated.Type
                        != RecordTextSpecialInfoAtomForPreservation
                    || generated.Version != 0 || generated.Instance != 0) {
                    return false;
                }
                var cursor = new LegacyPptTextPropertyCursor(generated,
                    "TextSpecialInfoAtom");
                using var rewritten = new MemoryStream();
                int covered = 0;
                while (!cursor.IsAtEnd) {
                    uint countValue = cursor.ReadUInt32();
                    if (countValue == 0 || countValue > int.MaxValue) {
                        return false;
                    }
                    covered = checked(covered + (int)countValue);
                    uint masks = cursor.ReadUInt32();
                    if ((masks & ~0x00000007U) != 0) return false;
                    ushort? spelling = (masks & 1U) != 0
                        ? cursor.ReadUInt16()
                        : null;
                    ushort? language = (masks & (1U << 1)) != 0
                        ? cursor.ReadUInt16()
                        : null;
                    ushort? alternative = (masks & (1U << 2)) != 0
                        ? cursor.ReadUInt16()
                        : null;
                    uint rewrittenMasks = masks;
                    if (forcePrimaryNoLanguage && !language.HasValue) {
                        rewrittenMasks |= 1U << 1;
                    }
                    if (forceAlternativeNoLanguage
                        && !alternative.HasValue) {
                        rewrittenMasks |= 1U << 2;
                    }
                    WriteUInt32ToStream(rewritten, countValue);
                    WriteUInt32ToStream(rewritten, rewrittenMasks);
                    if (spelling.HasValue) {
                        WriteUInt16ToStream(rewritten, spelling.Value);
                    }
                    if ((rewrittenMasks & (1U << 1)) != 0) {
                        WriteUInt16ToStream(rewritten, language ?? 0);
                    }
                    if ((rewrittenMasks & (1U << 2)) != 0) {
                        WriteUInt16ToStream(rewritten, alternative ?? 0);
                    }
                }
                if (covered != replacementCharacterCount) return false;
                record = BuildRecord(version: 0, instance: 0,
                    RecordTextSpecialInfoAtomForPreservation,
                    rewritten.ToArray());
                return true;
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentException) {
                record = replacement;
                return false;
            }
        }

        private static bool TryReadTextSpecialInfoDefaults(
            LegacyPptRecord record, int expectedCharacterCount,
            out bool forcePrimaryNoLanguage,
            out bool forceAlternativeNoLanguage) {
            forcePrimaryNoLanguage = false;
            forceAlternativeNoLanguage = false;
            if (record.Version != 0 || record.Instance != 0) return false;
            var cursor = new LegacyPptTextPropertyCursor(record,
                "TextSpecialInfoAtom");
            int covered = 0;
            bool allPrimaryNoLanguage = true;
            bool allAlternativeNoLanguage = true;
            bool sawRun = false;
            while (!cursor.IsAtEnd) {
                uint countValue = cursor.ReadUInt32();
                if (countValue == 0 || countValue > int.MaxValue) {
                    return false;
                }
                covered = checked(covered + (int)countValue);
                uint masks = cursor.ReadUInt32();
                if ((masks & ~0x00000007U) != 0) return false;
                if ((masks & 1U) != 0
                    && (cursor.ReadUInt16() & ~0x0003) != 0) {
                    return false;
                }
                allPrimaryNoLanguage &= (masks & (1U << 1)) != 0
                    && cursor.ReadUInt16() == 0;
                allAlternativeNoLanguage &= (masks & (1U << 2)) != 0
                    && cursor.ReadUInt16() == 0;
                sawRun = true;
            }
            if (!sawRun || covered != expectedCharacterCount) return false;
            forcePrimaryNoLanguage = allPrimaryNoLanguage;
            forceAlternativeNoLanguage = allAlternativeNoLanguage;
            return true;
        }

        private static void WriteUInt16ToStream(Stream stream,
            ushort value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32ToStream(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }

        private static bool TryRewriteTextFontCollection(
            LegacyPptPackage package, byte[]? currentDocumentBytes,
            LegacyPptWriter.LegacyPptWriterFontCatalog fonts,
            out byte[] bytes) {
            if (!package.PersistObjects.TryGetValue(
                    package.DocumentPersistId,
                    out LegacyPptPersistObject? persistObject)) {
                bytes = Array.Empty<byte>();
                return false;
            }
            byte[] source = currentDocumentBytes
                ?? persistObject.RecordBytes;
            try {
                LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(
                    source, 0, new LegacyPptImportOptions());
                if (fonts.TryRewriteCollection(document, out bytes)) {
                    return true;
                }
                if (fonts.HasPrototype || document.Version != 0x0F) {
                    bytes = source;
                    return false;
                }
                var children = document.Children.Select(child =>
                    child.CopyRecordBytes()).ToList();
                children.Add(fonts.BuildCollection());
                bytes = BuildRecord(document.Version, document.Instance,
                    document.Type, Concat(children));
                return true;
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentException) {
                bytes = source;
                return false;
            }
        }
    }
}
