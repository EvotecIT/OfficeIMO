using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// Appends a binary PowerPoint incremental edit for changes that can be represented without rebuilding
    /// or discarding the source persist graph. The original document stream remains an exact prefix.
    /// </summary>
    internal static class LegacyPptPreservingWriter {
        private const ushort RecordPersistDirectory = 0x1772;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtChildAnchor = 0xF00F;
        private const ushort OfficeArtClientAnchor = 0xF010;

        internal static bool CanWritePresentation(PowerPointPresentation presentation) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return TryBuildModifiedSlides(presentation, out _);
        }

        internal static bool TryWritePresentation(PowerPointPresentation presentation, out byte[] bytes) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            bytes = Array.Empty<byte>();
            if (!TryBuildModifiedSlides(presentation, out IReadOnlyDictionary<uint, byte[]> modifiedSlides)) {
                return false;
            }

            LegacyPptPackage package = presentation.LegacyPptPackage!;
            if (modifiedSlides.Count == 0) {
                bytes = package.CopyOriginalBytes();
                return true;
            }

            byte[] documentStream = AppendIncrementalEdit(package, modifiedSlides, out uint editOffset);
            byte[] currentUserStream = PatchCurrentEditOffset(package.CurrentUserStream, editOffset);
            bytes = OfficeCompoundFileWriter.Rewrite(package.CompoundFile,
                new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                    ["PowerPoint Document"] = documentStream,
                    ["Current User"] = currentUserStream
                });
            return true;
        }

        private static bool TryBuildModifiedSlides(PowerPointPresentation presentation,
            out IReadOnlyDictionary<uint, byte[]> modifiedSlides) {
            var rewritten = new Dictionary<uint, byte[]>();
            modifiedSlides = rewritten;
            LegacyPptPackage? package = presentation.LegacyPptPackage;
            LegacyPptProjectionMap? projectionMap = presentation.LegacyPptProjectionMap;
            if (package == null || projectionMap == null || !presentation.HasOnlyLegacyPptGeometryChanges
                || presentation.Slides.Count != projectionMap.Slides.Count) {
                return false;
            }

            try {
                foreach (PowerPointSlide slide in presentation.Slides) {
                    if (!projectionMap.TryGetSlide(slide, out LegacyPptSlideProjection? slideProjection)
                        || slideProjection == null
                        || !package.PersistObjects.TryGetValue(slideProjection.PersistId,
                            out LegacyPptPersistObject? persistObject)) {
                        return false;
                    }

                    PowerPointShape[] shapes = slide.Shapes.ToArray();
                    if (shapes.Length != slideProjection.Shapes.Count) return false;
                    var boundsByOfficeArtId = new Dictionary<uint, LegacyPptBounds>();
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
                        if (!BoundsEqual(bounds, shapeProjection.Bounds)) {
                            boundsByOfficeArtId.Add(shapeProjection.OfficeArtShapeId, bounds);
                        }
                    }
                    if (boundsByOfficeArtId.Count == 0) continue;

                    LegacyPptRecord slideRecord = LegacyPptRecordReader.ReadSingle(persistObject.RecordBytes, 0,
                        new LegacyPptImportOptions());
                    RecordRewrite result = RewriteRecord(slideRecord, boundsByOfficeArtId);
                    if (!result.Changed || result.PatchedShapeCount != boundsByOfficeArtId.Count) return false;
                    rewritten.Add(slideProjection.PersistId, result.Bytes);
                }
                return true;
            } catch (Exception exception) when (exception is InvalidDataException
                                                || exception is OverflowException
                                                || exception is ArgumentException) {
                rewritten.Clear();
                return false;
            }
        }

        private static RecordRewrite RewriteRecord(LegacyPptRecord record,
            IReadOnlyDictionary<uint, LegacyPptBounds> boundsByOfficeArtId) {
            if (record.Type == OfficeArtSpContainer) {
                LegacyPptRecord? fsp = record.Children.FirstOrDefault(child => child.Type == OfficeArtFsp);
                if (fsp != null && fsp.PayloadLength >= 4
                    && boundsByOfficeArtId.TryGetValue(fsp.ReadUInt32(0), out LegacyPptBounds bounds)) {
                    return RewriteShapeContainer(record, bounds);
                }
            }
            if (record.Version != 0x0F || record.Children.Count == 0) {
                return new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
            }

            var children = new List<byte[]>(record.Children.Count);
            bool changed = false;
            int patchedShapeCount = 0;
            foreach (LegacyPptRecord child in record.Children) {
                RecordRewrite childResult = RewriteRecord(child, boundsByOfficeArtId);
                children.Add(childResult.Bytes);
                changed |= childResult.Changed;
                patchedShapeCount = checked(patchedShapeCount + childResult.PatchedShapeCount);
            }
            return changed
                ? new RecordRewrite(BuildRecord(record.Version, record.Instance, record.Type, Concat(children)),
                    changed: true, patchedShapeCount)
                : new RecordRewrite(record.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
        }

        private static RecordRewrite RewriteShapeContainer(LegacyPptRecord shapeContainer, LegacyPptBounds bounds) {
            var children = new List<byte[]>(shapeContainer.Children.Count);
            bool patchedAnchor = false;
            foreach (LegacyPptRecord child in shapeContainer.Children) {
                if (!patchedAnchor && (child.Type == OfficeArtClientAnchor || child.Type == OfficeArtChildAnchor)) {
                    children.Add(BuildAnchor(child.Type, child.Instance, bounds));
                    patchedAnchor = true;
                } else {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (!patchedAnchor) {
                return new RecordRewrite(shapeContainer.CopyRecordBytes(), changed: false, patchedShapeCount: 0);
            }
            return new RecordRewrite(BuildRecord(shapeContainer.Version, shapeContainer.Instance,
                shapeContainer.Type, Concat(children)), changed: true, patchedShapeCount: 1);
        }

        private static byte[] AppendIncrementalEdit(LegacyPptPackage package,
            IReadOnlyDictionary<uint, byte[]> modifiedSlides, out uint editOffset) {
            using var output = new MemoryStream();
            output.Write(package.DocumentStream, 0, package.DocumentStream.Length);
            var offsets = new SortedDictionary<uint, uint>();
            foreach (KeyValuePair<uint, byte[]> slide in modifiedSlides.OrderBy(pair => pair.Key)) {
                offsets.Add(slide.Key, checked((uint)output.Position));
                output.Write(slide.Value, 0, slide.Value.Length);
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
            WriteUInt32(edit, 16, package.CurrentEditOffset);
            WriteUInt32(edit, 20, directoryOffset);
            WriteUInt32(edit, 24, package.DocumentPersistId);
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
            if (kind == LegacyPptShapeKind.TextBox) return shape is PowerPointTextBox;
            if (shape is not PowerPointAutoShape autoShape) return false;
            if (kind == LegacyPptShapeKind.Rectangle) return autoShape.ShapeType == A.ShapeTypeValues.Rectangle;
            if (kind == LegacyPptShapeKind.Ellipse) return autoShape.ShapeType == A.ShapeTypeValues.Ellipse;
            return kind == LegacyPptShapeKind.Line && autoShape.ShapeType == A.ShapeTypeValues.Line;
        }

        private static bool BoundsEqual(LegacyPptBounds left, LegacyPptBounds right) =>
            left.Left == right.Left && left.Top == right.Top && left.Width == right.Width && left.Height == right.Height;

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
    }
}
