namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    internal static void CollapseInlineMathRuns(
        OneNoteParagraph paragraph,
        int maximumNativeMathDepth = OneNoteReaderOptions.DefaultMaxPropertySetDepth) {
        if (!paragraph.Runs.Any(IsNativeMathRun)) return;
        var semanticRuns = new List<OneNoteTextRun>();
        int index = 0;
        while (index < paragraph.Runs.Count) {
            OneNoteTextRun source = paragraph.Runs[index];
            if (!IsNativeMathRun(source)) {
                semanticRuns.Add(source);
                index++;
                continue;
            }

            List<OneNoteTextRun> nativeRuns = ReadOneNativeMathExpression(paragraph.Runs, ref index);
            OfficeIMO.Drawing.OfficeMathExpression expression = OneNoteMathNativeCodec.Decode(nativeRuns, maximumNativeMathDepth);
            var semantic = new OneNoteTextRun {
                Hyperlink = nativeRuns[0].Hyperlink,
                HyperlinkProtected = nativeRuns[0].HyperlinkProtected
            };
            CopyTextStyle(nativeRuns[0].Style, semantic.Style);
            semantic.SetMathExpression(expression);
            semantic.PreservedMathExpression = expression;
            semantic.PreservedNativeMathRuns = nativeRuns.Select(OneNoteMathRunPreservation.Clone).ToArray();
            foreach (OneNoteTextRun native in nativeRuns) {
                foreach (OneNoteOpaqueProperty property in native.UnknownProperties) semantic.UnknownProperties.Add(property);
            }
            semanticRuns.Add(semantic);
        }

        paragraph.Runs.Clear();
        foreach (OneNoteTextRun run in semanticRuns) paragraph.Runs.Add(run);
    }

    private static bool IsNativeMathRun(OneNoteTextRun run) => run.Style.IsMath == true || run.MathDescriptor != null;

    private static List<OneNoteTextRun> ReadOneNativeMathExpression(IList<OneNoteTextRun> runs, ref int index) {
        var result = new List<OneNoteTextRun>();
        int objectDepth = 0;
        bool sawObject = false;
        while (index < runs.Count && IsNativeMathRun(runs[index])) {
            OneNoteTextRun run = runs[index++];
            result.Add(run);
            string text = run.Text ?? string.Empty;
            for (int characterIndex = 0; characterIndex < text.Length; characterIndex++) {
                if (text[characterIndex] == OneNoteMathNativeCodec.ObjectStart) {
                    sawObject = true;
                    objectDepth++;
                } else if (text[characterIndex] == OneNoteMathNativeCodec.ObjectEnd && objectDepth > 0) {
                    objectDepth--;
                }
            }
            if (!sawObject || objectDepth == 0) break;
        }
        return result;
    }

    private static void CopyTextStyle(OneNoteTextStyle source, OneNoteTextStyle destination) {
        destination.FontFamily = source.FontFamily;
        destination.FontSize = source.FontSize;
        destination.ColorArgb = source.ColorArgb;
        destination.HighlightColorArgb = source.HighlightColorArgb;
        destination.Bold = source.Bold;
        destination.Italic = source.Italic;
        destination.Underline = source.Underline;
        destination.Strikethrough = source.Strikethrough;
        destination.Superscript = source.Superscript;
        destination.Subscript = source.Subscript;
        destination.LanguageId = source.LanguageId;
        destination.IsMath = source.IsMath;
    }

    private static OneNoteMathInlineDescriptor? ReadMathDescriptor(OneNotePropertySet set) {
        uint? type = ReadUInt32(set, OneNoteSchema.MathInlineObjectType);
        if (!type.HasValue) return null;
        return new OneNoteMathInlineDescriptor {
            Type = type.Value,
            Count = ReadUInt32(set, OneNoteSchema.MathInlineObjectCount),
            Column = ReadByte(set, OneNoteSchema.MathInlineObjectColumn),
            Alignment = ReadByte(set, OneNoteSchema.MathInlineObjectAlignment),
            Character = ReadUInt16(set, OneNoteSchema.MathInlineObjectCharacter),
            Character1 = ReadUInt16(set, OneNoteSchema.MathInlineObjectCharacter1),
            Character2 = ReadUInt16(set, OneNoteSchema.MathInlineObjectCharacter2)
        };
    }

    private static byte? ReadByte(OneNotePropertySet? set, uint propertyId) {
        ulong? value = FindProperty(set, propertyId)?.ScalarValue;
        return value.HasValue ? (byte)value.Value : null;
    }

    private static ushort? ReadUInt16(OneNotePropertySet? set, uint propertyId) {
        ulong? value = FindProperty(set, propertyId)?.ScalarValue;
        return value.HasValue ? (ushort)value.Value : null;
    }
}
