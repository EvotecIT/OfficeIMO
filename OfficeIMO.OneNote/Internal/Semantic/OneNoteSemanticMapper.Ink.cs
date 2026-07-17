using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private sealed class InkRecognitionValue {
        internal InkRecognitionValue(IReadOnlyList<string> alternatives, uint? languageId) {
            Alternatives = alternatives;
            LanguageId = languageId;
        }
        internal IReadOnlyList<string> Alternatives { get; }
        internal uint? LanguageId { get; }
    }

    private static OneNoteInk BuildInk(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject container,
        OneNoteReaderOptions options,
        int depth,
        HashSet<OneNoteExtendedGuid> path) {
        var ink = new OneNoteInk { Id = container.Id, Layout = ReadLayout(container) };
        OneNoteExtendedGuid? inkDataId = GetReferences(container, OneNoteSchema.InkData).FirstOrDefault();
        if (inkDataId != null) {
            ink.InkDataObjectId = inkDataId;
            OneNoteRevisionStoreObject? inkData = space.GetObject(inkDataId);
            if (inkData != null && inkData.Jcid.Value == OneNoteSchema.JcidInkDataNode) {
                byte[]? boundingBox = ReadData(inkData, OneNoteSchema.InkBoundingBox);
                ink.PreservedInkBoundingBox = boundingBox == null ? null : (byte[])boundingBox.Clone();
                double scaleX = ReadFloat(container, OneNoteSchema.InkScalingX) ?? 1D;
                double scaleY = ReadFloat(container, OneNoteSchema.InkScalingY) ?? 1D;
                ink.PreservedInkScaleX = scaleX;
                ink.PreservedInkScaleY = scaleY;
                foreach (OneNoteExtendedGuid strokeId in GetReferences(inkData, OneNoteSchema.InkStrokes)) {
                    OneNoteRevisionStoreObject? strokeObject = space.GetObject(strokeId);
                    if (strokeObject == null || strokeObject.Jcid.Value != OneNoteSchema.JcidInkStrokeNode) {
                        ink.PreservedStrokeObjectIds.Add(strokeId);
                        continue;
                    }
                    OfficeInkStroke? stroke = DecodeInkStroke(
                        space,
                        strokeObject,
                        scaleX,
                        scaleY,
                        options,
                        out bool dimensionsContainUnsupportedPackets);
                    if (stroke == null) {
                        ink.PreservedStrokeObjectIds.Add(strokeId);
                        continue;
                    }
                    ink.Ink.Add(stroke);
                    ink.StrokeObjectIds[stroke] = strokeId;
                    if (dimensionsContainUnsupportedPackets) ink.PreservedNativeStrokeSnapshots[stroke] = stroke.Clone();
                    OneNoteExtendedGuid? propertyId = GetReferences(strokeObject, OneNoteSchema.InkStrokeProperties).FirstOrDefault();
                    if (propertyId != null) ink.StrokePropertyObjectIds[stroke] = propertyId;
                }
            }
            return ink;
        }

        if (depth + 1 >= options.MaxPropertySetDepth) return ink;
        foreach (OneNoteExtendedGuid childId in GetReferences(container, OneNoteSchema.ContentChildNodes)) {
            if (!path.Add(childId)) continue;
            try {
                OneNoteRevisionStoreObject? childObject = space.GetObject(childId);
                if (childObject == null || childObject.Jcid.Value != OneNoteSchema.JcidInkContainer) continue;
                ink.PreservedChildContainerIds.Add(childId);
                OneNoteInk child = BuildInk(space, childObject, options, depth + 1, path);
                double offsetX = child.Layout?.X ?? 0D;
                double offsetY = child.Layout?.Y ?? 0D;
                foreach (OfficeInkStroke childStroke in child.Strokes) {
                    OfficeInkStroke clone = childStroke.Clone();
                    OfficeTransform translation = OfficeTransform.Translate(offsetX, offsetY);
                    clone.Transform = clone.Transform.HasValue ? clone.Transform.Value.Then(translation) : translation;
                    ink.Ink.Add(clone);
                    ink.PreservedNestedStrokeSnapshots[clone] = clone.Clone();
                    if (child.StrokeObjectIds.TryGetValue(childStroke, out OneNoteExtendedGuid? strokeId)) ink.StrokeObjectIds[clone] = strokeId;
                    if (child.StrokePropertyObjectIds.TryGetValue(childStroke, out OneNoteExtendedGuid? propertyId)) ink.StrokePropertyObjectIds[clone] = propertyId;
                }
                foreach (OneNoteExtendedGuid preservedId in child.PreservedStrokeObjectIds) ink.PreservedStrokeObjectIds.Add(preservedId);
            } finally {
                path.Remove(childId);
            }
        }
        return ink;
    }

    private static OfficeInkStroke? DecodeInkStroke(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject source,
        double scaleX,
        double scaleY,
        OneNoteReaderOptions options,
        out bool dimensionsContainUnsupportedPackets) {
        dimensionsContainUnsupportedPackets = false;
        OneNoteExtendedGuid? propertiesId = GetReferences(source, OneNoteSchema.InkStrokeProperties).FirstOrDefault();
        OneNoteRevisionStoreObject? properties = propertiesId == null ? null : space.GetObject(propertiesId);
        if (properties == null || properties.Jcid.Value != OneNoteSchema.JcidStrokePropertiesNode) return null;
        byte[]? pathData = ReadData(source, OneNoteSchema.InkPath);
        if (pathData == null) return null;
        IReadOnlyList<OneNoteInkCodec.Dimension> dimensions = OneNoteInkCodec.DecodeDimensions(ReadData(properties, OneNoteSchema.InkDimensions));
        int xIndex = IndexOfDimension(dimensions, OneNoteInkCodec.XDimension);
        int yIndex = IndexOfDimension(dimensions, OneNoteInkCodec.YDimension);
        if (xIndex < 0 || yIndex < 0 || dimensions.Count == 0) return null;
        IReadOnlyList<long> values = OneNoteInkCodec.DecodeSignedVector(pathData, Math.Min(pathData.Length, (int)Math.Min(int.MaxValue, options.MaxInputBytes ?? int.MaxValue)));
        if (values.Count == 0 || values.Count % dimensions.Count != 0) return null;
        int pointCount = values.Count / dimensions.Count;
        int pressureIndex = IndexOfDimension(dimensions, OneNoteInkCodec.PressureDimension);
        dimensionsContainUnsupportedPackets = dimensions.Any(dimension =>
            dimension.Id != OneNoteInkCodec.XDimension &&
            dimension.Id != OneNoteInkCodec.YDimension &&
            dimension.Id != OneNoteInkCodec.PressureDimension);
        IReadOnlyList<long> xPackets = OneNoteInkCodec.DecodePacketValues(values, xIndex * pointCount, pointCount);
        IReadOnlyList<long> yPackets = OneNoteInkCodec.DecodePacketValues(values, yIndex * pointCount, pointCount);
        IReadOnlyList<long>? pressurePackets = pressureIndex < 0
            ? null
            : OneNoteInkCodec.DecodePacketValues(values, pressureIndex * pointCount, pointCount);
        var stroke = new OfficeInkStroke {
            Color = OneNoteInkCodec.DecodeColor(ReadUInt32(properties, OneNoteSchema.InkColor)),
            Width = Math.Max(0.000001D, (ReadFloat(properties, OneNoteSchema.InkWidth) ?? 1D) * Math.Abs(scaleX) / OneNoteInkCodec.NativeUnitsPerHalfInch),
            Height = Math.Max(0.000001D, (ReadFloat(properties, OneNoteSchema.InkHeight) ?? 1D) * Math.Abs(scaleY) / OneNoteInkCodec.NativeUnitsPerHalfInch),
            Opacity = 1D - (ReadByte(properties, OneNoteSchema.InkTransparency) ?? 0) / (double)byte.MaxValue,
            TipShape = (ReadByte(properties, OneNoteSchema.InkPenTip) ?? 0) == 0 ? OfficeInkTipShape.Ellipse : OfficeInkTipShape.Rectangle,
            FitToCurve = ReadBoolean(properties, OneNoteSchema.InkFitToCurve) ?? false,
            IgnorePressure = ReadBoolean(properties, OneNoteSchema.InkIgnorePressure) ?? false,
            Bias = DecodeInkBias(ReadByte(source, OneNoteSchema.InkBias)),
            LanguageId = ReadUInt16Value(source, OneNoteSchema.InkLanguageId)
        };
        for (int index = 0; index < pointCount; index++) {
            double x = xPackets[index] * scaleX / OneNoteInkCodec.NativeUnitsPerHalfInch;
            double y = yPackets[index] * scaleY / OneNoteInkCodec.NativeUnitsPerHalfInch;
            double? pressure = null;
            if (pressureIndex >= 0 && pressurePackets != null) {
                OneNoteInkCodec.Dimension pressureDimension = dimensions[pressureIndex];
                long raw = pressurePackets[index];
                long range = (long)pressureDimension.Upper - pressureDimension.Lower;
                if (range > 0) pressure = Math.Max(0D, Math.Min(1D, (raw - pressureDimension.Lower) / (double)range));
            }
            stroke.AddPoint(x, y, pressure);
        }
        return stroke;
    }

    private static IDictionary<OneNoteExtendedGuid, InkRecognitionValue> ReadInkRecognition(
        OneNoteMaterializedObjectSpace space,
        OneNoteRevisionStoreObject pageNode) {
        var map = new Dictionary<OneNoteExtendedGuid, InkRecognitionValue>();
        OneNoteExtendedGuid? rootId = GetReferences(pageNode, OneNoteSchema.PageRecognizedTextContainer).FirstOrDefault();
        if (rootId == null) return map;
        CollectInkRecognition(space, rootId, map, new HashSet<OneNoteExtendedGuid>(), 0);
        return map;
    }

    private static void CollectInkRecognition(
        OneNoteMaterializedObjectSpace space,
        OneNoteExtendedGuid id,
        IDictionary<OneNoteExtendedGuid, InkRecognitionValue> map,
        HashSet<OneNoteExtendedGuid> path,
        int depth) {
        if (depth > 8 || !path.Add(id)) return;
        try {
            OneNoteRevisionStoreObject? item = space.GetObject(id);
            if (item == null) return;
            if (item.Jcid.Value == OneNoteSchema.JcidRecognizedTextWord) {
                IReadOnlyList<string> alternatives = DecodeRecognitionAlternatives(ReadData(item, OneNoteSchema.RecognizedText));
                var value = new InkRecognitionValue(alternatives, ReadUInt16Value(item, OneNoteSchema.RecognizedTextLanguageId));
                byte[]? references = ReadData(item, OneNoteSchema.RecognizedTextStrokeReferences);
                if (references != null) {
                    for (int offset = 0; offset + 20 <= references.Length; offset += 20) {
                        uint allocation = OneNoteBinary.ReadUInt32(references, offset + 16);
                        // The leading GUID identifies the recognizer batch, not the ink object.
                        // Stroke allocations are resolved in the word node's object namespace.
                        map[new OneNoteExtendedGuid(id.Identifier, allocation, 17)] = value;
                    }
                }
                return;
            }
            foreach (OneNoteExtendedGuid childId in GetReferences(item, OneNoteSchema.RecognizedTextChildNodes)) {
                CollectInkRecognition(space, childId, map, path, depth + 1);
            }
        } finally {
            path.Remove(id);
        }
    }

    private static void ApplyInkRecognition(OneNotePage page, IDictionary<OneNoteExtendedGuid, InkRecognitionValue> recognition) {
        if (recognition.Count == 0) return;
        foreach (OneNoteElement element in OneNoteElementTraversal.Enumerate(page)) {
            if (!(element is OneNoteInk ink)) continue;
            foreach (OfficeInkStroke stroke in ink.Strokes) {
                if (!ink.StrokeObjectIds.TryGetValue(stroke, out OneNoteExtendedGuid? id) || !recognition.TryGetValue(id, out InkRecognitionValue? value)) continue;
                stroke.RecognizedText = value.Alternatives.FirstOrDefault();
                stroke.RecognitionAlternatives.Clear();
                foreach (string alternative in value.Alternatives) stroke.RecognitionAlternatives.Add(alternative);
                if (value.LanguageId.HasValue) stroke.LanguageId = value.LanguageId.Value;
                if (ink.PreservedNativeStrokeSnapshots.ContainsKey(stroke)) {
                    ink.PreservedNativeStrokeSnapshots[stroke] = stroke.Clone();
                }
                if (ink.PreservedNestedStrokeSnapshots.ContainsKey(stroke)) {
                    ink.PreservedNestedStrokeSnapshots[stroke] = stroke.Clone();
                }
            }
        }
    }

    private static IReadOnlyList<string> DecodeRecognitionAlternatives(byte[]? data) {
        if (data == null || data.Length < 2) return Array.Empty<string>();
        string text = System.Text.Encoding.Unicode.GetString(data, 0, data.Length - data.Length % 2);
        return text.Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
    }

    private static int IndexOfDimension(IReadOnlyList<OneNoteInkCodec.Dimension> dimensions, Guid id) {
        for (int index = 0; index < dimensions.Count; index++) if (dimensions[index].Id == id) return index;
        return -1;
    }

    private static ushort? ReadUInt16Value(OneNoteRevisionStoreObject? item, uint propertyId) {
        ulong? value = FindProperty(item?.PropertySet, propertyId)?.ScalarValue;
        return value.HasValue ? (ushort)value.Value : null;
    }

    private static OfficeInkBias DecodeInkBias(byte? value) {
        switch (value) {
            case 0: return OfficeInkBias.Handwriting;
            case 1: return OfficeInkBias.Drawing;
            default: return OfficeInkBias.Both;
        }
    }
}
