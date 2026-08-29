using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionStructureInspector {
    internal static PdfPrintProductionStructureEvidence Inspect(
        PdfReadDocument document,
        System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        int validBoxes = 0;
        int invalidBoxes = 0;
        foreach (PdfReadPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasValidProductionBoxes(page.GetGeometry())) validBoxes++;
            else invalidBoxes++;
        }

        ReachableFontInspection fontInspection = InspectReachableFonts(document, cancellationToken);
        HashSet<PdfDictionary> fontDictionaries = fontInspection.Fonts;

        int unembedded = 0;
        int uninspectable = fontInspection.UninspectableContextCount;
        foreach (PdfDictionary font in fontDictionaries) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                fontInspection.SelectedType3CharacterCodes.TryGetValue(font, out HashSet<int>? selectedType3CharacterCodes);
                if (!HasEmbeddedFontProgram(
                        font,
                        document.Objects,
                        document.ReadOptions.Limits.MaxDecodedStreamBytes,
                        document.ReadOptions.Limits.MaxObjectNestingDepth,
                        selectedType3CharacterCodes)) unembedded++;
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is NotSupportedException ||
                exception is PdfReadLimitException) {
                uninspectable++;
            }
        }

        return new PdfPrintProductionStructureEvidence(
            document.Pages.Count,
            validBoxes,
            invalidBoxes,
            fontDictionaries.Count,
            unembedded,
            uninspectable);
    }

    private static bool HasValidProductionBoxes(PdfPageGeometry geometry) {
        PdfPageBox? media = geometry.MediaBox;
        PdfPageBox? trim = geometry.TrimBox;
        PdfPageBox? bleed = geometry.BleedBox;
        PdfPageBox? art = geometry.ArtBox;
        if (media == null || (trim == null) == (art == null)) return false;
        PdfPageBox productionBoundary = trim ?? art!;
        if (!Contains(media, productionBoundary)) return false;
        return bleed == null ||
            (Contains(media, bleed) && Contains(bleed, productionBoundary));
    }

    private static bool Contains(PdfPageBox outer, PdfPageBox inner) {
        const double tolerance = 0.0001D;
        return inner.Left >= outer.Left - tolerance &&
            inner.Bottom >= outer.Bottom - tolerance &&
            inner.Right <= outer.Right + tolerance &&
            inner.Top <= outer.Top + tolerance;
    }

    private static PdfObject? ResolveObject(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        int depth,
        int maximumObjectDepth,
        out int resolvedDepth) {
        var visited = new HashSet<(int ObjectNumber, int Generation)>();
        PdfObject? resolved = value;
        resolvedDepth = depth;
        while (resolved is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation)) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject indirect)) return null;
            resolvedDepth++;
            if (resolvedDepth > maximumObjectDepth) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.ObjectNestingDepth,
                    maximumObjectDepth,
                    resolvedDepth);
            }
            resolved = indirect.Value;
        }
        return resolved;
    }

    private static bool HasEmbeddedFontProgram(
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth,
        HashSet<int>? selectedType3CharacterCodes) {
        string? subtype = ResolveName(
            font.Items.TryGetValue("Subtype", out PdfObject? subtypeObject) ? subtypeObject : null,
            objects,
            maximumObjectDepth);
        if (string.Equals(subtype, "Type3", StringComparison.Ordinal)) {
            return HasValidType3FontProgram(
                font,
                objects,
                maxDecodedStreamBytes,
                maximumObjectDepth,
                selectedType3CharacterCodes);
        }

        PdfDictionary fontWithDescriptor = font;
        string? programFontSubtype = subtype;
        if (string.Equals(subtype, "Type0", StringComparison.Ordinal)) {
            if (!font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) ||
                ResolveObject(objects, descendantsObject, 0, maximumObjectDepth, out _) is not PdfArray descendants ||
                descendants.Items.Count != 1 ||
                ResolveObject(objects, descendants.Items[0], 0, maximumObjectDepth, out _) is not PdfDictionary descendant) {
                return false;
            }
            string? descendantSubtype = ResolveName(
                descendant.Items.TryGetValue("Subtype", out PdfObject? descendantSubtypeObject)
                    ? descendantSubtypeObject
                    : null,
                objects,
                maximumObjectDepth);
            if (!string.Equals(descendantSubtype, "CIDFontType0", StringComparison.Ordinal) &&
                !string.Equals(descendantSubtype, "CIDFontType2", StringComparison.Ordinal)) {
                return false;
            }
            fontWithDescriptor = descendant;
            programFontSubtype = descendantSubtype;
        }

        if (!fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject) ||
            ResolveObject(objects, descriptorObject, 0, maximumObjectDepth, out _) is not PdfDictionary descriptor) return false;

        return HasValidFontDescriptorAndProgram(
            programFontSubtype,
            descriptor,
            objects,
            maxDecodedStreamBytes,
            maximumObjectDepth);
    }

    private static bool HasValidType3FontProgram(
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth,
        HashSet<int>? selectedCharacterCodes) {
        if (!string.Equals(ResolveName(
                font.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null,
                objects,
                maximumObjectDepth), "Font", StringComparison.Ordinal) ||
            !TryReadFiniteNumberArray(font, "FontBBox", 4, objects, maximumObjectDepth, out _) ||
            !TryReadFiniteNumberArray(font, "FontMatrix", 6, objects, maximumObjectDepth, out double[] matrix) ||
            Math.Abs((matrix[0] * matrix[3]) - (matrix[1] * matrix[2])) <= 0.000000000001D ||
            !TryReadInteger(font, "FirstChar", objects, maximumObjectDepth, 0, 255, out int firstCharacter) ||
            !TryReadInteger(font, "LastChar", objects, maximumObjectDepth, firstCharacter, 255, out int lastCharacter) ||
            !font.Items.TryGetValue("Widths", out PdfObject? widthsObject) ||
            ResolveObject(objects, widthsObject, 0, maximumObjectDepth, out _) is not PdfArray widths ||
            widths.Items.Count != lastCharacter - firstCharacter + 1 ||
            !AllFiniteNumbers(widths, objects, maximumObjectDepth) ||
            !font.Items.TryGetValue("CharProcs", out PdfObject? charProcsObject) ||
            ResolveObject(objects, charProcsObject, 0, maximumObjectDepth, out _) is not PdfDictionary charProcs ||
            charProcs.Items.Count is < 1 or > 256 ||
            !PdfPrintProductionColorInspector.TryGetType3GlyphNames(
                font,
                objects,
                maximumObjectDepth,
                out Dictionary<int, string> glyphNames)) {
            return false;
        }

        if (font.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
            ResolveObject(objects, resourcesObject, 0, maximumObjectDepth, out _) is not (PdfNull or PdfDictionary)) {
            return false;
        }

        if (selectedCharacterCodes == null || selectedCharacterCodes.Count == 0) return true;
        foreach (int characterCode in selectedCharacterCodes) {
            if (characterCode < firstCharacter || characterCode > lastCharacter ||
                !glyphNames.TryGetValue(characterCode, out string? glyphName) ||
                !charProcs.Items.TryGetValue(glyphName, out PdfObject? charProcObject) ||
                ResolveObject(objects, charProcObject, 0, maximumObjectDepth, out _) is not PdfStream charProc ||
                !HasValidType3MetricsOperator(charProc, objects, maxDecodedStreamBytes, maximumObjectDepth)) {
                return false;
            }
        }
        return true;
    }

    private static bool HasValidType3MetricsOperator(
        PdfStream charProc,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth) {
        if (StreamDecoder.GetUnsupportedFilters(charProc.Dictionary, objects).Count != 0 ||
            !StreamDecoder.TryDecode(
                charProc.Dictionary,
                charProc.Data,
                maxDecodedStreamBytes,
                out byte[] decoded,
                objects)) return false;

        PdfContentOperation? firstOperation = null;
        try {
            PdfContentStreamInterpreter.InterpretUntil(
                PdfEncoding.Latin1GetString(decoded),
                maxOperations: 1,
                operation => {
                    firstOperation = operation;
                    return false;
                },
                maxNestingDepth: maximumObjectDepth);
        } catch (InvalidDataException) {
            return false;
        } catch (PdfReadLimitException) {
            throw;
        }

        if (firstOperation is not PdfContentOperation metricsOperation || metricsOperation.HasInvalidOperands) return false;
        int expectedOperands = string.Equals(metricsOperation.Name, "d0", StringComparison.Ordinal)
            ? 2
            : string.Equals(metricsOperation.Name, "d1", StringComparison.Ordinal)
                ? 6
                : 0;
        return expectedOperands > 0 &&
            metricsOperation.Operands.Count == expectedOperands &&
            metricsOperation.Operands.All(static operand => operand is double value && !double.IsNaN(value) && !double.IsInfinity(value));
    }

    private static bool TryReadFiniteNumberArray(
        PdfDictionary dictionary,
        string key,
        int expectedCount,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out double[] values) {
        values = Array.Empty<double>();
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(objects, value, 0, maximumObjectDepth, out _) is not PdfArray array ||
            array.Items.Count != expectedCount) return false;
        values = new double[expectedCount];
        for (int index = 0; index < expectedCount; index++) {
            if (ResolveObject(objects, array.Items[index], 0, maximumObjectDepth, out _) is not PdfNumber number ||
                double.IsNaN(number.Value) || double.IsInfinity(number.Value)) {
                values = Array.Empty<double>();
                return false;
            }
            values[index] = number.Value;
        }
        return true;
    }

    private static bool TryReadInteger(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int minimum,
        int maximum,
        out int value) {
        value = 0;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? item) ||
            ResolveObject(objects, item, 0, maximumObjectDepth, out _) is not PdfNumber number ||
            double.IsNaN(number.Value) || double.IsInfinity(number.Value) ||
            number.Value != Math.Truncate(number.Value) ||
            number.Value < minimum || number.Value > maximum) return false;
        value = (int)number.Value;
        return true;
    }

    private static bool AllFiniteNumbers(
        PdfArray values,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        values.Items.All(value =>
            ResolveObject(objects, value, 0, maximumObjectDepth, out _) is PdfNumber number &&
            !double.IsNaN(number.Value) &&
            !double.IsInfinity(number.Value));

    private static bool HasReadableFontStream(
        string? fontSubtype,
        PdfDictionary descriptor,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth) {
        if (!descriptor.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(objects, value, 0, maximumObjectDepth, out _) is not PdfStream stream ||
            StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) return false;
        return StreamDecoder.TryDecode(
            stream.Dictionary,
            stream.Data,
            maxDecodedStreamBytes,
            out byte[] decoded,
            objects) && IsValidFontProgram(fontSubtype, key, stream, decoded, objects, maximumObjectDepth);
    }

    private static bool HasValidFontDescriptorAndProgram(
        string? fontSubtype,
        PdfDictionary descriptor,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedStreamBytes,
        int maximumObjectDepth) {
        if (!string.Equals(
                ResolveName(
                    descriptor.Items.TryGetValue("Type", out PdfObject? typeObject) ? typeObject : null,
                    objects,
                    maximumObjectDepth),
                "FontDescriptor",
                StringComparison.Ordinal) ||
            ResolveName(
                descriptor.Items.TryGetValue("FontName", out PdfObject? fontNameObject) ? fontNameObject : null,
                objects,
                maximumObjectDepth) is not { Length: > 0 } ||
            !TryReadInteger(descriptor, "Flags", objects, maximumObjectDepth, 0, int.MaxValue, out _) ||
            !TryReadFiniteNumberArray(descriptor, "FontBBox", 4, objects, maximumObjectDepth, out double[] fontBox) ||
            fontBox[2] < fontBox[0] ||
            fontBox[3] < fontBox[1] ||
            !TryReadFiniteNumber(descriptor, "ItalicAngle", objects, maximumObjectDepth, out _) ||
            !TryReadFiniteNumber(descriptor, "Ascent", objects, maximumObjectDepth, out _) ||
            !TryReadFiniteNumber(descriptor, "Descent", objects, maximumObjectDepth, out _) ||
            !TryReadFiniteNumber(descriptor, "CapHeight", objects, maximumObjectDepth, out _) ||
            !TryReadFiniteNumber(descriptor, "StemV", objects, maximumObjectDepth, out double stemV) ||
            stemV <= 0D) return false;

        string? programKey = null;
        foreach (string key in new[] { "FontFile", "FontFile2", "FontFile3" }) {
            if (!descriptor.Items.TryGetValue(key, out PdfObject? candidate)) continue;
            PdfObject? resolved = ResolveObject(objects, candidate, 0, maximumObjectDepth, out _);
            if (resolved is PdfNull) continue;
            if (resolved is not PdfStream || programKey != null) return false;
            programKey = key;
        }

        return programKey != null && HasReadableFontStream(
            fontSubtype,
            descriptor,
            programKey,
            objects,
            maxDecodedStreamBytes,
            maximumObjectDepth);
    }

    private static bool TryReadFiniteNumber(
        PdfDictionary dictionary,
        string key,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        out double value) {
        value = 0D;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? candidate) ||
            ResolveObject(objects, candidate, 0, maximumObjectDepth, out _) is not PdfNumber number ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value)) return false;
        value = number.Value;
        return true;
    }

    private static string? ResolveName(
        PdfObject? value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) =>
        value != null && ResolveObject(objects, value, 0, maximumObjectDepth, out _) is PdfName name ? name.Name : null;
}
