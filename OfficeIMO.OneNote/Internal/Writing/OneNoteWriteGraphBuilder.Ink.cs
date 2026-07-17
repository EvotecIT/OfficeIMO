using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

internal sealed partial class OneNoteWriteGraphBuilder {
    private OneNoteExtendedGuid BuildInk(OneNoteWriteObjectSpace space, OneNoteInk ink, uint lastModifiedTime) {
        if (CanPreserveNestedInkContainer(ink, out IReadOnlyList<OfficeInkStroke>? authoredStrokes)) {
            var retainedChildren = new List<OneNoteExtendedGuid>(ink.PreservedChildContainerIds);
            if (authoredStrokes.Count > 0) {
                var authoredInk = new OneNoteInk();
                foreach (OfficeInkStroke stroke in authoredStrokes) authoredInk.Ink.Add(stroke);
                OneNoteExtendedGuid authoredContainerId = BuildInk(space, authoredInk, lastModifiedTime);
                retainedChildren.Add(authoredContainerId);
                foreach (OfficeInkStroke stroke in authoredStrokes) {
                    if (authoredInk.StrokeObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? strokeId)) ink.StrokeObjectIds[stroke] = strokeId;
                    if (authoredInk.StrokePropertyObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? propertyId)) ink.StrokePropertyObjectIds[stroke] = propertyId;
                }
            }
            OneNoteExtendedGuid retainedContainerId = IdOrNew(ink.Id);
            ink.Id = retainedContainerId;
            var retainedContainerProperties = LayoutProperties(ink.Layout);
            retainedContainerProperties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
            retainedContainerProperties.Add(ObjectReferences(OneNoteSchema.ContentChildNodes, retainedChildren));
            space.Objects.Add(new OneNoteWriteObject(retainedContainerId, OneNoteSchema.JcidInkContainer, retainedContainerProperties));
            return retainedContainerId;
        }

        if (ink.PreservedChildContainerIds.Count > 0 && ink.PreservedStrokeObjectIds.Count > 0) {
            throw new OneNoteFormatException(
                "ONENOTE_WRITE_OPAQUE_NESTED_INK_EDIT",
                "Nested ink containing undecodable native strokes cannot be flattened after an existing stroke is edited.");
        }

        bool retainsNativeStrokes = ink.PreservedStrokeObjectIds.Count > 0 || ink.Strokes.Any(stroke => CanReuseNativeStroke(ink, stroke));
        double scaleX = retainsNativeStrokes ? ValidInkScale(ink.PreservedInkScaleX) : 1D;
        double scaleY = retainsNativeStrokes ? ValidInkScale(ink.PreservedInkScaleY) : 1D;
        var strokeIds = new List<OneNoteExtendedGuid>(ink.PreservedStrokeObjectIds);
        var nativePoints = new List<OfficePoint>();
        foreach (OfficeInkStroke stroke in ink.Strokes) {
            if (stroke.Points.Count == 0) continue;
            if (CanReuseNativeStroke(ink, stroke) && ink.StrokeObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? preservedStrokeId)) {
                stroke.ValidateForOneNote();
                AddNativePoints(stroke, nativePoints, scaleX, scaleY);
                strokeIds.Add(preservedStrokeId);
                continue;
            }
            strokeIds.Add(BuildInkStroke(space, ink, stroke, nativePoints, scaleX, scaleY));
        }

        OneNoteExtendedGuid dataId = IdOrNew(ink.InkDataObjectId);
        ink.InkDataObjectId = dataId;
        var dataProperties = new List<OneNoteWriteProperty>();
        if (strokeIds.Count > 0) dataProperties.Add(ObjectReferences(OneNoteSchema.InkStrokes, strokeIds));
        byte[]? boundingBox = InkBoundingBox(ink, nativePoints, retainsNativeStrokes);
        if (boundingBox != null) dataProperties.Add(Data(OneNoteSchema.InkBoundingBox, boundingBox));
        space.Objects.Add(new OneNoteWriteObject(dataId, OneNoteSchema.JcidInkDataNode, dataProperties));

        OneNoteExtendedGuid containerId = IdOrNew(ink.Id);
        ink.Id = containerId;
        var containerProperties = LayoutProperties(ink.Layout);
        containerProperties.Insert(0, Scalar(OneNoteSchema.LastModifiedTime, lastModifiedTime));
        containerProperties.Add(ObjectReferences(OneNoteSchema.InkData, dataId));
        // Preservation merging retains unknown source properties. Emit canonical scale explicitly so
        // ordinary re-encoded strokes are not scaled twice; retain the source scale only when raw
        // packet dimensions force at least one native stroke to be reused.
        containerProperties.Add(Float(OneNoteSchema.InkScalingX, scaleX));
        containerProperties.Add(Float(OneNoteSchema.InkScalingY, scaleY));
        space.Objects.Add(new OneNoteWriteObject(containerId, OneNoteSchema.JcidInkContainer, containerProperties));
        return containerId;
    }

    private bool CanPreserveNestedInkContainer(OneNoteInk ink, out IReadOnlyList<OfficeInkStroke> authoredStrokes) {
        authoredStrokes = Array.Empty<OfficeInkStroke>();
        if (ink.PreservedChildContainerIds.Count == 0 ||
            ink.Id == null || _activeSourceSpace?.GetObject(ink.Id) == null ||
            !ink.PreservedChildContainerIds.All(id => _activeSourceSpace.GetObject(id) != null)) return false;
        foreach (KeyValuePair<OfficeInkStroke, OfficeInkStroke> retained in ink.PreservedNestedStrokeSnapshots) {
            if (!ink.Strokes.Contains(retained.Key) || !NativeStrokePayloadEquals(retained.Key, retained.Value)) return false;
        }
        authoredStrokes = ink.Strokes.Where(stroke => !ink.PreservedNestedStrokeSnapshots.ContainsKey(stroke)).ToArray();
        return true;
    }

    private bool CanReuseNativeStroke(OneNoteInk ink, OfficeInkStroke stroke) =>
        ink.PreservedNativeStrokeSnapshots.TryGetValue(stroke, out OfficeInkStroke? snapshot) &&
        NativeStrokePayloadEquals(stroke, snapshot) &&
        ink.StrokeObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? preservedStrokeId) &&
        _activeSourceSpace?.GetObject(preservedStrokeId) != null;

    private OneNoteExtendedGuid BuildInkStroke(
        OneNoteWriteObjectSpace space,
        OneNoteInk owner,
        OfficeInkStroke stroke,
        IList<OfficePoint> allNativePoints,
        double containerScaleX,
        double containerScaleY) {
        stroke.ValidateForOneNote();
        OfficeTransform transform = stroke.Transform ?? OfficeTransform.Identity;
        var xValues = new List<long>(stroke.Points.Count);
        var yValues = new List<long>(stroke.Points.Count);
        var pressureValues = new List<long>(stroke.Points.Count);
        bool hasPressure = stroke.Points.Any(point => point.Pressure.HasValue);
        for (int index = 0; index < stroke.Points.Count; index++) {
            OfficeInkPoint point = stroke.Points[index].Transform(transform);
            int x = OneNoteInkCodec.ToNativeCoordinate(point.X / containerScaleX);
            int y = OneNoteInkCodec.ToNativeCoordinate(point.Y / containerScaleY);
            xValues.Add(x);
            yValues.Add(y);
            allNativePoints.Add(new OfficePoint(x, y));
            if (hasPressure) pressureValues.Add((long)Math.Round(Math.Max(0D, Math.Min(1D, point.Pressure ?? 1D)) * 32767D));
        }
        var pathValues = new List<long>(xValues.Count + yValues.Count + pressureValues.Count);
        pathValues.AddRange(OneNoteInkCodec.EncodePacketValues(xValues));
        pathValues.AddRange(OneNoteInkCodec.EncodePacketValues(yValues));
        pathValues.AddRange(OneNoteInkCodec.EncodePacketValues(pressureValues));

        (double transformedTipWidth, double transformedTipHeight) = stroke.GetTransformedTipDimensions();
        OneNoteExtendedGuid propertyId = owner.StrokePropertyObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? retainedPropertyId)
            ? retainedPropertyId
            : _ids.New();
        owner.StrokePropertyObjectIds[stroke] = propertyId;
        var propertyValues = new List<OneNoteWriteProperty> {
            Data(OneNoteSchema.InkDimensions, OneNoteInkCodec.EncodeDimensions(hasPressure)),
            Scalar(OneNoteSchema.InkColor, OneNoteInkCodec.EncodeColor(stroke.Color)),
            Float(OneNoteSchema.InkHeight, transformedTipHeight * OneNoteInkCodec.NativeUnitsPerHalfInch / Math.Abs(containerScaleY)),
            Float(OneNoteSchema.InkWidth, transformedTipWidth * OneNoteInkCodec.NativeUnitsPerHalfInch / Math.Abs(containerScaleX)),
            Boolean(OneNoteSchema.InkAntialised, true),
            Boolean(OneNoteSchema.InkFitToCurve, stroke.FitToCurve),
            Boolean(OneNoteSchema.InkIgnorePressure, stroke.IgnorePressure),
            Scalar(OneNoteSchema.InkPenTip, stroke.TipShape == OfficeInkTipShape.Rectangle ? 1UL : 0UL),
            Scalar(OneNoteSchema.InkTransparency, (ulong)Math.Round((1D - OfficeInkRenderer.GetEffectiveOpacity(stroke)) * byte.MaxValue))
        };
        space.Objects.Add(new OneNoteWriteObject(propertyId, OneNoteSchema.JcidStrokePropertiesNode, propertyValues));

        OneNoteExtendedGuid strokeId = owner.StrokeObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? retainedStrokeId)
            ? retainedStrokeId
            : _ids.New();
        owner.StrokeObjectIds[stroke] = strokeId;
        var strokeProperties = new List<OneNoteWriteProperty> {
            Data(OneNoteSchema.InkPath, OneNoteInkCodec.EncodeSignedVector(pathValues)),
            Scalar(OneNoteSchema.InkBias, EncodeInkBias(stroke.Bias)),
            ObjectReferences(OneNoteSchema.InkStrokeProperties, propertyId)
        };
        if (stroke.LanguageId.HasValue) strokeProperties.Insert(2, Scalar(OneNoteSchema.InkLanguageId, Math.Min(ushort.MaxValue, stroke.LanguageId.Value)));
        space.Objects.Add(new OneNoteWriteObject(strokeId, OneNoteSchema.JcidInkStrokeNode, strokeProperties));
        return strokeId;
    }

    private static void AddNativePoints(
        OfficeInkStroke stroke,
        ICollection<OfficePoint> output,
        double containerScaleX,
        double containerScaleY) {
        OfficeTransform transform = stroke.Transform ?? OfficeTransform.Identity;
        for (int index = 0; index < stroke.Points.Count; index++) {
            OfficeInkPoint point = stroke.Points[index].Transform(transform);
            output.Add(new OfficePoint(
                OneNoteInkCodec.ToNativeCoordinate(point.X / containerScaleX),
                OneNoteInkCodec.ToNativeCoordinate(point.Y / containerScaleY)));
        }
    }

    private static bool NativeStrokePayloadEquals(OfficeInkStroke left, OfficeInkStroke right) {
        if (left.Color != right.Color || left.Width != right.Width || left.Height != right.Height ||
            left.Opacity != right.Opacity || left.TipShape != right.TipShape || left.Bias != right.Bias ||
            left.FitToCurve != right.FitToCurve || left.IgnorePressure != right.IgnorePressure ||
            left.IsHighlighter != right.IsHighlighter || !Nullable.Equals(left.Transform, right.Transform) ||
            left.LanguageId != right.LanguageId ||
            left.Points.Count != right.Points.Count) return false;
        for (int index = 0; index < left.Points.Count; index++) {
            if (!left.Points[index].Equals(right.Points[index])) return false;
        }
        return true;
    }

    private OneNoteExtendedGuid? BuildInkRecognition(
        OneNoteWriteObjectSpace space,
        OneNotePage page,
        OneNoteMaterializedObjectSpace? sourceSpace) {
        var wordIds = new List<OneNoteExtendedGuid>();
        Guid recognizerBatchId = Guid.NewGuid();
        OneNoteInk[] inks = OneNoteElementTraversal.Enumerate(page).OfType<OneNoteInk>().ToArray();
        foreach (OneNoteInk ink in inks) {
            foreach (OfficeInkStroke stroke in ink.Strokes) {
                IReadOnlyList<string> alternatives = RecognitionAlternatives(stroke);
                if (alternatives.Count == 0 || !ink.StrokeObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? strokeId)) continue;

                OneNoteExtendedGuid wordId = NewRecognitionWordId(space, sourceSpace, strokeId.Identifier);
                var properties = new List<OneNoteWriteProperty> {
                    Data(OneNoteSchema.RecognizedText, EncodeRecognitionAlternatives(alternatives)),
                    Data(OneNoteSchema.RecognizedTextStrokeReferences, EncodeRecognitionStrokeReferences(recognizerBatchId, ink.Id, strokeId))
                };
                if (stroke.LanguageId.HasValue) {
                    properties.Add(Scalar(OneNoteSchema.RecognizedTextLanguageId, Math.Min(ushort.MaxValue, stroke.LanguageId.Value)));
                }
                space.Objects.Add(new OneNoteWriteObject(wordId, OneNoteSchema.JcidRecognizedTextWord, properties));
                wordIds.Add(wordId);
            }
        }
        OneNoteExtendedGuid? authoredLineId = null;
        if (wordIds.Count > 0) {
            OneNoteExtendedGuid blockId = _ids.New();
            space.Objects.Add(new OneNoteWriteObject(
                blockId,
                OneNoteSchema.JcidRecognizedTextBlock,
                new[] { ObjectReferences(OneNoteSchema.RecognizedTextChildNodes, wordIds) }));
            authoredLineId = _ids.New();
            space.Objects.Add(new OneNoteWriteObject(
                authoredLineId,
                OneNoteSchema.JcidRecognizedTextLine,
                new[] { ObjectReferences(OneNoteSchema.RecognizedTextChildNodes, blockId) }));
        }

        var opaqueStrokeIds = new HashSet<OneNoteExtendedGuid>(inks.SelectMany(ink => ink.PreservedStrokeObjectIds));
        OneNoteExtendedGuid? preservedRootId = BuildPreservedRecognitionTree(
            space,
            page,
            sourceSpace,
            opaqueStrokeIds,
            authoredLineId);
        if (preservedRootId != null) return preservedRootId;
        if (authoredLineId == null) return null;

        OneNoteExtendedGuid rootId = _ids.New();
        space.Objects.Add(new OneNoteWriteObject(
            rootId,
            OneNoteSchema.JcidRecognizedTextRoot,
            new[] { ObjectReferences(OneNoteSchema.RecognizedTextChildNodes, authoredLineId) }));
        return rootId;
    }

    private static OneNoteExtendedGuid? BuildPreservedRecognitionTree(
        OneNoteWriteObjectSpace outputSpace,
        OneNotePage page,
        OneNoteMaterializedObjectSpace? sourceSpace,
        IReadOnlyCollection<OneNoteExtendedGuid> opaqueStrokeIds,
        OneNoteExtendedGuid? appendedLineId) {
        if (sourceSpace == null || page.PreservationIds.PageNodeId == null || opaqueStrokeIds.Count == 0) {
            return null;
        }

        OneNoteRevisionStoreObject? pageNode = sourceSpace.GetObject(page.PreservationIds.PageNodeId);
        OneNoteExtendedGuid? rootId = OneNoteSemanticMapper
            .GetReferences(pageNode, OneNoteSchema.PageRecognizedTextContainer)
            .FirstOrDefault();
        if (rootId == null) return null;

        return PreserveRecognitionNode(
            outputSpace,
            sourceSpace,
            rootId,
            opaqueStrokeIds,
            new HashSet<OneNoteExtendedGuid>(),
            0,
            appendedLineId);
    }

    private static OneNoteExtendedGuid? PreserveRecognitionNode(
        OneNoteWriteObjectSpace outputSpace,
        OneNoteMaterializedObjectSpace sourceSpace,
        OneNoteExtendedGuid id,
        IReadOnlyCollection<OneNoteExtendedGuid> opaqueStrokeIds,
        ISet<OneNoteExtendedGuid> path,
        int depth,
        OneNoteExtendedGuid? appendedRootChildId = null) {
        if (depth > 8 || !path.Add(id)) return null;
        try {
            OneNoteRevisionStoreObject? item = sourceSpace.GetObject(id);
            if (item == null) return null;
            if (item.Jcid.Value == OneNoteSchema.JcidRecognizedTextWord) {
                byte[]? references = OneNoteSemanticMapper.ReadData(item, OneNoteSchema.RecognizedTextStrokeReferences);
                if (references == null) return null;
                for (int offset = 0; offset + 20 <= references.Length; offset += 20) {
                    uint allocation = OneNoteBinary.ReadUInt32(references, offset + 16);
                    if (opaqueStrokeIds.Contains(new OneNoteExtendedGuid(id.Identifier, allocation, 17))) {
                        return id;
                    }
                }
                return null;
            }

            var retainedChildren = new List<OneNoteExtendedGuid>();
            foreach (OneNoteExtendedGuid childId in OneNoteSemanticMapper.GetReferences(item, OneNoteSchema.RecognizedTextChildNodes)) {
                OneNoteExtendedGuid? retainedChild = PreserveRecognitionNode(
                    outputSpace,
                    sourceSpace,
                    childId,
                    opaqueStrokeIds,
                    path,
                    depth + 1);
                if (retainedChild != null) retainedChildren.Add(retainedChild);
            }
            if (appendedRootChildId != null) retainedChildren.Add(appendedRootChildId);
            if (retainedChildren.Count == 0) return null;

            outputSpace.Objects.Add(new OneNoteWriteObject(
                id,
                item.Jcid.Value,
                new[] { ObjectReferences(OneNoteSchema.RecognizedTextChildNodes, retainedChildren) }));
            return id;
        } finally {
            path.Remove(id);
        }
    }

    private static OneNoteExtendedGuid NewRecognitionWordId(
        OneNoteWriteObjectSpace space,
        OneNoteMaterializedObjectSpace? sourceSpace,
        Guid identifier) {
        var used = new HashSet<uint>(space.Objects.Where(item => item.Id.Identifier == identifier).Select(item => item.Id.Value));
        if (sourceSpace != null) {
            foreach (OneNoteRevisionStoreObject item in sourceSpace.Objects) {
                if (item.Id.Identifier == identifier) used.Add(item.Id.Value);
            }
        }
        for (uint allocation = 1; allocation <= byte.MaxValue; allocation++) {
            if (!used.Contains(allocation)) return new OneNoteExtendedGuid(identifier, allocation, 17);
        }
        throw new OneNoteFormatException(
            "ONENOTE_WRITE_RECOGNITION_ID_LIMIT",
            "A handwriting-recognition word cannot be allocated in its ink stroke namespace because all CompactID allocations are in use.");
    }

    private static IReadOnlyList<string> RecognitionAlternatives(OfficeInkStroke stroke) {
        var values = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        if (!string.IsNullOrWhiteSpace(stroke.RecognizedText) && seen.Add(stroke.RecognizedText!)) values.Add(stroke.RecognizedText!);
        foreach (string alternative in stroke.RecognitionAlternatives) {
            if (!string.IsNullOrWhiteSpace(alternative) && seen.Add(alternative)) values.Add(alternative);
        }
        return values;
    }

    private static byte[] EncodeRecognitionAlternatives(IReadOnlyList<string> alternatives) =>
        System.Text.Encoding.Unicode.GetBytes(string.Join("\0", alternatives) + "\0");

    private static byte[] EncodeRecognitionStrokeReferences(
        Guid recognizerBatchId,
        OneNoteExtendedGuid? inkContainerId,
        OneNoteExtendedGuid strokeId) {
        using (var stream = new MemoryStream(40)) {
            if (inkContainerId != null && inkContainerId.Identifier == strokeId.Identifier) {
                OneNoteDesktopBinary.WriteExtendedGuid(stream, new OneNoteExtendedGuid(recognizerBatchId, inkContainerId.Value, 17));
            }
            OneNoteDesktopBinary.WriteExtendedGuid(stream, new OneNoteExtendedGuid(recognizerBatchId, strokeId.Value, 17));
            return stream.ToArray();
        }
    }

    private static byte[]? InkBoundingBox(OneNoteInk ink, IReadOnlyList<OfficePoint> points, bool retainsNativeStrokes) {
        byte[]? source = ink.PreservedInkBoundingBox;
        bool sourceBoundsAreValid = TryReadInkBoundingBox(source, out int left, out int top, out int right, out int bottom);
        bool hasSourceBounds = retainsNativeStrokes && sourceBoundsAreValid;
        if (ink.PreservedStrokeObjectIds.Count > 0 && !hasSourceBounds) {
            throw new OneNoteFormatException(
                "ONENOTE_WRITE_OPAQUE_INK_BOUNDS",
                "Ink containing undecodable native strokes cannot be written without its complete native bounding box.");
        }
        if (points.Count == 0) return hasSourceBounds ? (byte[])source!.Clone() : null;

        int pointLeft = NativeCoordinate(points[0].X);
        int pointTop = NativeCoordinate(points[0].Y);
        int pointRight = pointLeft;
        int pointBottom = pointTop;
        if (!hasSourceBounds) {
            left = pointLeft;
            top = pointTop;
            right = pointRight;
            bottom = pointBottom;
        } else {
            left = Math.Min(left, pointLeft);
            top = Math.Min(top, pointTop);
            right = Math.Max(right, pointRight);
            bottom = Math.Max(bottom, pointBottom);
        }
        for (int index = 1; index < points.Count; index++) {
            int x = NativeCoordinate(points[index].X);
            int y = NativeCoordinate(points[index].Y);
            left = Math.Min(left, x);
            top = Math.Min(top, y);
            right = Math.Max(right, x);
            bottom = Math.Max(bottom, y);
        }
        var data = new byte[16];
        WriteInt32LittleEndian(data, 0, left);
        WriteInt32LittleEndian(data, 4, top);
        WriteInt32LittleEndian(data, 8, right);
        WriteInt32LittleEndian(data, 12, bottom);
        return data;
    }

    private static bool TryReadInkBoundingBox(byte[]? data, out int left, out int top, out int right, out int bottom) {
        left = top = right = bottom = 0;
        if (data == null || data.Length != 16) return false;
        left = unchecked((int)OneNoteBinary.ReadUInt32(data, 0));
        top = unchecked((int)OneNoteBinary.ReadUInt32(data, 4));
        right = unchecked((int)OneNoteBinary.ReadUInt32(data, 8));
        bottom = unchecked((int)OneNoteBinary.ReadUInt32(data, 12));
        return left <= right && top <= bottom;
    }

    private static int NativeCoordinate(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < int.MinValue || value > int.MaxValue) {
            throw new OneNoteFormatException("ONENOTE_WRITE_INK_BOUNDS", "Ink coordinates must fit the native 32-bit bounding-box range.");
        }
        return (int)value;
    }

    private static void WriteInt32LittleEndian(byte[] data, int offset, int value) {
        uint unsigned = unchecked((uint)value);
        data[offset] = (byte)unsigned;
        data[offset + 1] = (byte)(unsigned >> 8);
        data[offset + 2] = (byte)(unsigned >> 16);
        data[offset + 3] = (byte)(unsigned >> 24);
    }

    private static ulong EncodeInkBias(OfficeInkBias bias) {
        switch (bias) {
            case OfficeInkBias.Handwriting: return 0UL;
            case OfficeInkBias.Drawing: return 1UL;
            default: return 2UL;
        }
    }

    private static double ValidInkScale(double value) =>
        double.IsNaN(value) || double.IsInfinity(value) || Math.Abs(value) < 0.000001D ? 1D : value;
}

internal static class OneNoteInkWriterValidationExtensions {
    internal static void ValidateForOneNote(this OfficeInkStroke stroke) {
        if (double.IsNaN(stroke.Width) || double.IsInfinity(stroke.Width) || stroke.Width <= 0D ||
            double.IsNaN(stroke.Height) || double.IsInfinity(stroke.Height) || stroke.Height <= 0D) {
            throw new OneNoteFormatException("ONENOTE_WRITE_INK_STYLE", "Ink stroke dimensions must be finite and positive.");
        }
        if (double.IsNaN(stroke.Opacity) || double.IsInfinity(stroke.Opacity) || stroke.Opacity < 0D || stroke.Opacity > 1D) {
            throw new OneNoteFormatException("ONENOTE_WRITE_INK_STYLE", "Ink stroke opacity must be from 0 through 1.");
        }
    }
}
