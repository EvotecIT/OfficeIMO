using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordStyleTextPropAtomForPreservation = 0x0FA1;
        private const ushort RecordTextRulerAtomForPreservation = 0x0FA6;

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
                    == RecordTextRulerAtomForPreservation) > 1) {
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
                    if (content.RulerRecord != null) {
                        children.Add(content.RulerRecord);
                    }
                    wroteText = true;
                    continue;
                }
                if (child.Type == RecordStyleTextPropAtomForPreservation
                    || child.Type == RecordTextRulerAtomForPreservation) {
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
            bytes = BuildRecord(textbox.Version, textbox.Instance,
                textbox.Type, Concat(children));
            return true;
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
