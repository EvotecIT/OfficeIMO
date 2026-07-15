using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// Appends a binary PowerPoint incremental edit for changes that can be represented without rebuilding
    /// or discarding the source persist graph. The original document stream remains an exact prefix.
    /// </summary>
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordPersistDirectory = 0x1772;
        private const ushort RecordSlidePersistAtom = 0x03F3;
        private const ushort RecordSlideAtom = 0x03EF;
        private const ushort RecordSlideShowSlideInfoAtom = 0x03F9;
        private const ushort RecordNamedShows = 0x0410;
        private const ushort RecordSlideListWithText = 0x0FF0;
        private const ushort RecordTextHeader = 0x0F9F;
        private const ushort RecordTextChars = 0x0FA0;
        private const ushort RecordTextBytes = 0x0FA8;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtDgg = 0xF006;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtClientTextbox = 0xF00D;
        private const ushort OfficeArtChildAnchor = 0xF00F;
        private const ushort OfficeArtClientAnchor = 0xF010;

        internal static bool CanWritePresentation(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return TryBuildModifiedPersistObjects(presentation, out _, out _);
        }

        internal static bool TryWritePresentation(PowerPointPresentation presentation, out byte[] bytes) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            bytes = Array.Empty<byte>();
            if (!TryBuildModifiedPersistObjects(presentation,
                    out IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects,
                    out IReadOnlyList<uint> currentSlideIds)) {
                return false;
            }

            LegacyPptPackage package = presentation.LegacyPptPackage!;
            if (modifiedPersistObjects.Count == 0) {
                bytes = package.CopyOriginalBytes();
                return true;
            }

            byte[] documentStream = AppendIncrementalEdit(package, modifiedPersistObjects, currentSlideIds,
                out uint editOffset);
            byte[] currentUserStream = PatchCurrentEditOffset(package.CurrentUserStream, editOffset);
            bytes = OfficeCompoundFileWriter.Rewrite(package.CompoundFile,
                new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = documentStream,
                    ["Current User"] = currentUserStream
                });
            return true;
        }

        private static bool TryBuildModifiedPersistObjects(PowerPointPresentation presentation,
            out IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects,
            out IReadOnlyList<uint> currentSlideIds) {
            var rewritten = new Dictionary<uint, byte[]>();
            var slideIds = new List<uint>(presentation.Slides.Count);
            modifiedPersistObjects = rewritten;
            currentSlideIds = slideIds;
            LegacyPptPackage? package = presentation.LegacyPptPackage;
            LegacyPptProjectionMap? projectionMap = presentation.LegacyPptProjectionMap;
            if (package == null || projectionMap == null || !presentation.HasOnlyLegacyPptProjectedShapeChanges
                || presentation.Slides.Count > 4082) {
                return false;
            }

            try {
                var currentSlideOrder = new List<LegacyPptSlideProjection>(presentation.Slides.Count);
                var addedSlides = new List<PowerPointSlide>();
                bool encounteredAddedSlide = false;
                foreach (PowerPointSlide slide in presentation.Slides) {
                    if (!projectionMap.TryGetSlide(slide, out LegacyPptSlideProjection? slideProjection)
                        || slideProjection == null) {
                        if (!LegacyPptWritePreflight.CanWriteSlideLosslessly(slide)) return false;
                        encounteredAddedSlide = true;
                        addedSlides.Add(slide);
                        continue;
                    }
                    if (encounteredAddedSlide
                        || !package.PersistObjects.TryGetValue(slideProjection.PersistId,
                            out LegacyPptPersistObject? persistObject)
                        || persistObject == null) {
                        return false;
                    }
                    currentSlideOrder.Add(slideProjection);
                    slideIds.Add(slideProjection.SlideId);

                    PowerPointShape[] shapes = slide.Shapes.ToArray();
                    if (shapes.Length != slideProjection.Shapes.Count) return false;
                    var editsByOfficeArtId = new Dictionary<uint, ProjectedShapeEdit>();
                    foreach (PowerPointShape shape in shapes) {
                        uint? openXmlShapeId = shape.Id;
                        if (!openXmlShapeId.HasValue
                            || !slideProjection.TryGetShape(openXmlShapeId.Value,
                                out LegacyPptShapeProjection? shapeProjection)
                            || shapeProjection == null
                            || !MatchesProjectedKind(shape, shapeProjection.Kind)) {
                            return false;
                        }
                        LegacyPptBounds bounds = GetBounds(shape);
                        LegacyPptBounds? changedBounds = BoundsEqual(bounds, shapeProjection.Bounds)
                            ? null
                            : bounds;
                        string? changedText = null;
                        if (shape is PowerPointTextBox textBox) {
                            if (!HasOnlyPlainProjectedText(textBox)) return false;
                            string currentText = NormalizeLogicalText(textBox.Text);
                            if (!string.Equals(currentText, NormalizeLogicalText(shapeProjection.Text),
                                    StringComparison.Ordinal)) {
                                changedText = currentText;
                            }
                        }
                        if (changedBounds.HasValue || changedText != null) {
                            editsByOfficeArtId.Add(shapeProjection.OfficeArtShapeId,
                                new ProjectedShapeEdit(changedBounds, shapeProjection.Text, changedText));
                        }
                    }
                    bool? hidden = slide.Hidden == slideProjection.Hidden ? null : slide.Hidden;
                    if (editsByOfficeArtId.Count == 0 && !hidden.HasValue) continue;

                    LegacyPptRecord slideRecord = LegacyPptRecordReader.ReadSingle(persistObject.RecordBytes, 0,
                        new LegacyPptImportOptions());
                    if (!TryRewriteSlide(slideRecord, editsByOfficeArtId, hidden, out RecordRewrite result)
                        || !result.Changed || result.PatchedShapeCount != editsByOfficeArtId.Count) return false;
                    rewritten.Add(slideProjection.PersistId, result.Bytes);
                }
                bool originalTopologyChanged = !currentSlideOrder.Select(slide => slide.PersistId)
                    .SequenceEqual(projectionMap.Slides.Select(slide => slide.PersistId));
                if (addedSlides.Count > 0) {
                    if (originalTopologyChanged || currentSlideOrder.Count != projectionMap.Slides.Count
                        || !TryAppendNewSlides(package, projectionMap, addedSlides, rewritten,
                            out IReadOnlyList<uint> addedSlideIds)) {
                        return false;
                    }
                    slideIds.AddRange(addedSlideIds);
                } else if (originalTopologyChanged) {
                    if (!TryRewriteDocumentSlideOrder(package, projectionMap, currentSlideOrder,
                            out byte[] documentRecord)) {
                        return false;
                    }
                    rewritten.Add(package.DocumentPersistId, documentRecord);
                }
                return true;
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentException) {
                rewritten.Clear();
                return false;
            }
        }

        private static bool TryRewriteSlide(LegacyPptRecord slideRecord,
            IReadOnlyDictionary<uint, ProjectedShapeEdit> editsByOfficeArtId, bool? hidden,
            out RecordRewrite result) {
            bool hasSlideShowInfo = slideRecord.Children.Any(child => child.Type == RecordSlideShowSlideInfoAtom);
            bool patchedHidden = !hidden.HasValue;
            bool changed = false;
            int patchedShapeCount = 0;
            var children = new List<byte[]>(slideRecord.Children.Count + 1);
            foreach (LegacyPptRecord child in slideRecord.Children) {
                if (hidden.HasValue && child.Type == RecordSlideShowSlideInfoAtom) {
                    children.Add(PatchHiddenState(child.CopyRecordBytes(), hidden.Value));
                    patchedHidden = true;
                    changed = true;
                } else {
                    RecordRewrite childResult = RewriteRecord(child, editsByOfficeArtId);
                    children.Add(childResult.Bytes);
                    changed |= childResult.Changed;
                    patchedShapeCount = checked(patchedShapeCount + childResult.PatchedShapeCount);
                }
                if (hidden == true && !hasSlideShowInfo && child.Type == RecordSlideAtom) {
                    children.Add(BuildSlideShowInfo(hidden: true));
                    patchedHidden = true;
                    changed = true;
                }
            }
            if (!patchedHidden) {
                result = new RecordRewrite(slideRecord.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
                return false;
            }
            result = changed
                ? new RecordRewrite(BuildRecord(slideRecord.Version, slideRecord.Instance, slideRecord.Type,
                    Concat(children)), changed: true, patchedShapeCount)
                : new RecordRewrite(slideRecord.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
            return true;
        }

        private static RecordRewrite RewriteRecord(LegacyPptRecord record,
            IReadOnlyDictionary<uint, ProjectedShapeEdit> editsByOfficeArtId) {
            if (record.Type == OfficeArtSpContainer) {
                LegacyPptRecord? fsp = record.Children.FirstOrDefault(child => child.Type == OfficeArtFsp);
                if (fsp != null && fsp.PayloadLength >= 4
                    && editsByOfficeArtId.TryGetValue(fsp.ReadUInt32(0), out ProjectedShapeEdit? edit)
                    && edit != null) {
                    return TryRewriteShapeContainer(record, edit, out byte[] rewrittenShape)
                        ? new RecordRewrite(rewrittenShape, changed: true, patchedShapeCount: 1)
                        : new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
                }
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                return new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
            }

            var children = new List<byte[]>(record.Children.Count);
            bool changed = false;
            int patchedShapeCount = 0;
            foreach (LegacyPptRecord child in record.Children) {
                RecordRewrite childResult = RewriteRecord(child, editsByOfficeArtId);
                children.Add(childResult.Bytes);
                changed |= childResult.Changed;
                patchedShapeCount = checked(patchedShapeCount + childResult.PatchedShapeCount);
            }
            return changed
                ? new RecordRewrite(BuildRecord(record.Version, record.Instance, record.Type, Concat(children)),
                    changed: true, patchedShapeCount)
                : new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
        }

        private static bool TryRewriteShapeContainer(LegacyPptRecord shapeContainer, ProjectedShapeEdit edit,
            out byte[] bytes) {
            var children = new List<byte[]>(shapeContainer.Children.Count);
            bool patchedAnchor = !edit.Bounds.HasValue;
            bool patchedText = edit.Text == null;
            foreach (LegacyPptRecord child in shapeContainer.Children) {
                if (!patchedAnchor && edit.Bounds.HasValue
                    && (child.Type == OfficeArtClientAnchor || child.Type == OfficeArtChildAnchor)) {
                    children.Add(BuildAnchor(child.Type, child.Instance, edit.Bounds.Value));
                    patchedAnchor = true;
                } else if (!patchedText && child.Type == OfficeArtClientTextbox
                           && TryRewriteTextBox(child, edit.OriginalText, edit.Text!, out byte[] textbox)) {
                    children.Add(textbox);
                    patchedText = true;
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!patchedAnchor || !patchedText) {
                bytes = shapeContainer.CopyRecordBytes();
                return false;
            }
            bytes = BuildRecord(shapeContainer.Version, shapeContainer.Instance,
                shapeContainer.Type, Concat(children));
            return true;
        }

        private static bool TryRewriteTextBox(LegacyPptRecord textbox, string originalText, string replacementText,
            out byte[] bytes) {
            LegacyPptRecord[] textRecords = textbox.DescendantsAndSelf().Where(record =>
                record.Type == RecordTextChars || record.Type == RecordTextBytes).ToArray();
            if (textRecords.Length != 1
                || !TryBuildTextRecord(textbox, textRecords[0], originalText, replacementText,
                    out byte[] replacementRecord)) {
                bytes = textbox.CopyRecordBytes();
                return false;
            }
            return TryReplaceDescendant(textbox, textRecords[0].Offset, replacementRecord, out bytes);
        }

        private static bool TryBuildTextRecord(LegacyPptRecord textbox, LegacyPptRecord textRecord,
            string originalText, string replacementText, out byte[] bytes) {
            string raw = textRecord.Type == RecordTextChars
                ? textRecord.ReadUtf16Text()
                : textRecord.ReadLowByteUnicodeText();
            int contentLength = raw.Length;
            while (contentLength > 0 && raw[contentLength - 1] == '\0') contentLength--;
            while (contentLength > 0 && raw[contentLength - 1] == '\r') contentLength--;
            string decodedOriginal = NormalizeLogicalText(raw.Substring(0, contentLength));
            if (!string.Equals(decodedOriginal, NormalizeLogicalText(originalText), StringComparison.Ordinal)) {
                bytes = textRecord.CopyRecordBytes();
                return false;
            }

            string normalizedReplacement = NormalizeLogicalText(replacementText);
            if (normalizedReplacement.Length != contentLength && !IsStructurallyPlainTextBox(textbox)) {
                bytes = textRecord.CopyRecordBytes();
                return false;
            }
            string binaryReplacement = normalizedReplacement.Replace("\n", "\r") + raw.Substring(contentLength);
            byte[] payload;
            if (textRecord.Type == RecordTextChars) {
                payload = Encoding.Unicode.GetBytes(binaryReplacement);
            } else {
                if (binaryReplacement.Any(character => character > byte.MaxValue)) {
                    bytes = textRecord.CopyRecordBytes();
                    return false;
                }
                payload = binaryReplacement.Select(character => unchecked((byte)character)).ToArray();
            }
            bytes = BuildRecord(textRecord.Version, textRecord.Instance, textRecord.Type, payload);
            return true;
        }

        private static bool IsStructurallyPlainTextBox(LegacyPptRecord textbox) => textbox.Children.All(child =>
            child.Type == RecordTextHeader || child.Type == RecordTextChars || child.Type == RecordTextBytes);

        private static bool TryReplaceDescendant(LegacyPptRecord record, int targetOffset, byte[] replacement,
            out byte[] bytes) {
            if (record.Offset == targetOffset) {
                bytes = replacement;
                return true;
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                bytes = record.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(record.Children.Count);
            bool changed = false;
            foreach (LegacyPptRecord child in record.Children) {
                if (!changed && TryReplaceDescendant(child, targetOffset, replacement, out byte[] rewrittenChild)) {
                    children.Add(rewrittenChild);
                    changed = true;
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            bytes = changed
                ? BuildRecord(record.Version, record.Instance, record.Type, Concat(children))
                : record.CopyRecordBytes();
            return changed;
        }

        private static byte[] AppendIncrementalEdit(LegacyPptPackage package,
            IReadOnlyDictionary<uint, byte[]> modifiedPersistObjects, IReadOnlyList<uint> currentSlideIds,
            out uint editOffset) {
            using var output = new MemoryStream();
            output.Write(package.DocumentStream, 0, package.DocumentStream.Length);
            var offsets = new SortedDictionary<uint, uint>();
            foreach (KeyValuePair<uint, byte[]> persistObject in modifiedPersistObjects.OrderBy(pair => pair.Key)) {
                offsets.Add(persistObject.Key, checked((uint)output.Position));
                output.Write(persistObject.Value, 0, persistObject.Value.Length);
            }

            uint directoryOffset = checked((uint)output.Position);
            byte[] directory = BuildPersistDirectory(offsets);
            output.Write(directory, 0, directory.Length);

            editOffset = checked((uint)output.Position);
            LegacyPptRecord currentEdit = LegacyPptRecordReader.ReadSingle(package.DocumentStream,
                checked((int)package.CurrentEditOffset), new LegacyPptImportOptions());
            byte[] edit = currentEdit.CopyRecordBytes();
            if (currentEdit.PayloadLength < 20) {
                throw new InvalidDataException("The current UserEditAtom is too short for an incremental edit.");
            }
            uint lastViewedSlideId = ReadUInt32(edit, 8);
            if (lastViewedSlideId != 0 && !currentSlideIds.Contains(lastViewedSlideId)) {
                WriteUInt32(edit, 8, currentSlideIds.Count == 0 ? 0U : currentSlideIds[currentSlideIds.Count - 1]);
            }
            WriteUInt32(edit, 16, package.CurrentEditOffset);
            WriteUInt32(edit, 20, directoryOffset);
            WriteUInt32(edit, 24, package.DocumentPersistId);
            if (currentEdit.PayloadLength >= 24 && offsets.Count > 0) {
                WriteUInt32(edit, 28, Math.Max(currentEdit.ReadUInt32(20), offsets.Keys.Max()));
            }
            output.Write(edit, 0, edit.Length);
            return output.ToArray();
        }

        private static byte[] BuildPersistDirectory(IReadOnlyDictionary<uint, uint> offsets) {
            var payload = new List<byte>();
            KeyValuePair<uint, uint>[] entries = offsets.OrderBy(pair => pair.Key).ToArray();
            for (int index = 0; index < entries.Length;) {
                int count = 1;
                while (index + count < entries.Length && count < 0x0FFF
                       && entries[index + count].Key == entries[index].Key + unchecked((uint)count)) {
                    count++;
                }
                AppendUInt32(payload, (unchecked((uint)count) << 20) | entries[index].Key);
                for (int item = 0; item < count; item++) AppendUInt32(payload, entries[index + item].Value);
                index += count;
            }
            return BuildRecord(version: 0, instance: 0, RecordPersistDirectory, payload.ToArray());
        }

        private static byte[] PatchCurrentEditOffset(byte[] currentUserStream, uint editOffset) {
            byte[] patched = (byte[])currentUserStream.Clone();
            LegacyPptRecord currentUser = LegacyPptRecordReader.ReadSingle(patched, 0, new LegacyPptImportOptions());
            if (currentUser.PayloadLength < 12) {
                throw new InvalidDataException("The CurrentUserAtom is too short for its current-edit pointer.");
            }
            WriteUInt32(patched, 16, editOffset);
            return patched;
        }

        private static byte[] PatchHiddenState(byte[] slideShowInfo, bool hidden) {
            if (slideShowInfo.Length < 19) {
                throw new InvalidDataException("The slide-show information atom is too short for its flags.");
            }
            slideShowInfo[18] = hidden
                ? unchecked((byte)(slideShowInfo[18] | 0x04))
                : unchecked((byte)(slideShowInfo[18] & ~0x04));
            return slideShowInfo;
        }

        private static byte[] BuildSlideShowInfo(bool hidden) {
            var payload = new byte[16];
            payload[10] = hidden ? (byte)0x05 : (byte)0x01;
            return BuildRecord(version: 0, instance: 0, RecordSlideShowSlideInfoAtom, payload);
        }

        private static LegacyPptBounds GetBounds(PowerPointShape shape) {
            int left = ToMasterUnits(shape.Left);
            int top = ToMasterUnits(shape.Top);
            int width = Math.Max(0, ToMasterUnits(shape.Width));
            int height = Math.Max(0, ToMasterUnits(shape.Height));
            return new LegacyPptBounds(left, top, width, height);
        }

        private static byte[] BuildAnchor(ushort type, ushort instance, LegacyPptBounds bounds) {
            int right = checked(bounds.Left + bounds.Width);
            int bottom = checked(bounds.Top + bounds.Height);
            if (FitsInt16(bounds.Left) && FitsInt16(bounds.Top) && FitsInt16(right) && FitsInt16(bottom)) {
                var payload = new byte[8];
                WriteInt16(payload, 0, unchecked((short)bounds.Top));
                WriteInt16(payload, 2, unchecked((short)bounds.Left));
                WriteInt16(payload, 4, unchecked((short)right));
                WriteInt16(payload, 6, unchecked((short)bottom));
                return BuildRecord(version: 0, instance, type, payload);
            }
            var largePayload = new byte[16];
            WriteInt32(largePayload, 0, bounds.Top);
            WriteInt32(largePayload, 4, bounds.Left);
            WriteInt32(largePayload, 8, right);
            WriteInt32(largePayload, 12, bottom);
            return BuildRecord(version: 0, instance, type, largePayload);
        }

        private static bool MatchesProjectedKind(PowerPointShape shape, LegacyPptShapeKind kind) {
            if (kind == LegacyPptShapeKind.TextBox) {
                return shape is PowerPointTextBox;
            }
            if (kind == LegacyPptShapeKind.Picture) return shape is PowerPointPicture;
            if (kind == LegacyPptShapeKind.Connector) return shape is PowerPointConnectionShape;
            if (kind == LegacyPptShapeKind.Group) return shape is PowerPointGroupShape;
            if (shape is not PowerPointAutoShape autoShape) return false;
            if (kind == LegacyPptShapeKind.AutoShape) return autoShape.ShapeType.HasValue;
            if (kind == LegacyPptShapeKind.Rectangle) return autoShape.ShapeType == A.ShapeTypeValues.Rectangle;
            if (kind == LegacyPptShapeKind.Ellipse) return autoShape.ShapeType == A.ShapeTypeValues.Ellipse;
            return kind == LegacyPptShapeKind.Line && autoShape.ShapeType == A.ShapeTypeValues.Line;
        }

        private static bool HasOnlyPlainProjectedText(PowerPointTextBox textBox) {
            P.Shape? shape = textBox.Element as P.Shape;
            if (shape?.TextBody == null) return true;
            if (shape.TextBody.Descendants<A.Field>().Any() || shape.TextBody.Descendants<A.Break>().Any()) {
                return false;
            }
            return !shape.TextBody.Descendants<A.RunProperties>().Any(properties =>
                       properties.HasAttributes || properties.HasChildren)
                && !shape.TextBody.Descendants<A.ParagraphProperties>().Any(properties =>
                    properties.HasAttributes || properties.HasChildren)
                && !shape.TextBody.Descendants<A.EndParagraphRunProperties>().Any(properties =>
                    properties.HasAttributes || properties.HasChildren);
        }

        private static bool BoundsEqual(LegacyPptBounds left, LegacyPptBounds right) =>
            left.Left == right.Left && left.Top == right.Top && left.Width == right.Width && left.Height == right.Height;

        private static string NormalizeLogicalText(string value) => (value ?? string.Empty)
            .Replace("\r\n", "\n").Replace("\r", "\n");

        private static int ToMasterUnits(long emus) => checked((int)Math.Round(
            emus / 1587.5d, MidpointRounding.AwayFromZero));

        private static bool FitsInt16(int value) => value >= short.MinValue && value <= short.MaxValue;

        private static byte[] BuildRecord(byte version, ushort instance, ushort type, byte[] payload) {
            var bytes = new byte[checked(8 + payload.Length)];
            WriteUInt16(bytes, 0, unchecked((ushort)((instance << 4) | version)));
            WriteUInt16(bytes, 2, type);
            WriteUInt32(bytes, 4, unchecked((uint)payload.Length));
            Buffer.BlockCopy(payload, 0, bytes, 8, payload.Length);
            return bytes;
        }

        private static byte[] Concat(IEnumerable<byte[]> records) {
            byte[][] values = records.ToArray();
            var result = new byte[values.Sum(record => record.Length)];
            int offset = 0;
            foreach (byte[] record in values) {
                Buffer.BlockCopy(record, 0, result, offset, record.Length);
                offset += record.Length;
            }
            return result;
        }

        private static void AppendUInt32(ICollection<byte> bytes, uint value) {
            bytes.Add(unchecked((byte)value));
            bytes.Add(unchecked((byte)(value >> 8)));
            bytes.Add(unchecked((byte)(value >> 16)));
            bytes.Add(unchecked((byte)(value >> 24)));
        }

        private static void WriteInt16(byte[] bytes, int offset, short value) =>
            WriteUInt16(bytes, offset, unchecked((ushort)value));

        private static uint ReadUInt32(byte[] bytes, int offset) => unchecked((uint)(bytes[offset]
            | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24));

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) =>
            WriteUInt32(bytes, offset, unchecked((uint)value));

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = unchecked((byte)value);
            bytes[offset + 1] = unchecked((byte)(value >> 8));
            bytes[offset + 2] = unchecked((byte)(value >> 16));
            bytes[offset + 3] = unchecked((byte)(value >> 24));
        }

        private readonly struct RecordRewrite {
            internal RecordRewrite(byte[] bytes, bool changed, int patchedShapeCount) {
                Bytes = bytes;
                Changed = changed;
                PatchedShapeCount = patchedShapeCount;
            }

            internal byte[] Bytes { get; }

            internal bool Changed { get; }

            internal int PatchedShapeCount { get; }
        }

        private sealed class ProjectedShapeEdit {
            internal ProjectedShapeEdit(LegacyPptBounds? bounds, string originalText, string? text) {
                Bounds = bounds;
                OriginalText = originalText ?? string.Empty;
                Text = text;
            }

            internal LegacyPptBounds? Bounds { get; }

            internal string OriginalText { get; }

            internal string? Text { get; }
        }
    }
}
