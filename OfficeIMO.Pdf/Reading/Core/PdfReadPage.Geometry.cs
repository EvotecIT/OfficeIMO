namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    /// <summary>Reads page boundary boxes and page-level presentation metadata.</summary>
    public PdfPageGeometry GetGeometry() {
        PdfPageTransition? transition = ReadTransition(GetDirectValue("Trans"));
        bool hasMetadata = HasDirectKey("Metadata");
        int? metadataObjectNumber = GetDirectValue("Metadata") is PdfReference metadataReference
            ? metadataReference.ObjectNumber
            : null;

        return new PdfPageGeometry(
            TryReadPageBox("MediaBox", out PdfPageBox? mediaBox) ? mediaBox : null,
            TryReadPageBox("CropBox", out PdfPageBox? cropBox) ? cropBox : null,
            TryReadPageBox("BleedBox", out PdfPageBox? bleedBox) ? bleedBox : null,
            TryReadPageBox("TrimBox", out PdfPageBox? trimBox) ? trimBox : null,
            TryReadPageBox("ArtBox", out PdfPageBox? artBox) ? artBox : null,
            TryReadDirectPositiveNumber("UserUnit"),
            TryReadDirectName("Tabs"),
            TryReadDirectPositiveOrZeroNumber("Dur"),
            transition,
            hasMetadata,
            metadataObjectNumber,
            HasDirectKey("PieceInfo"));
    }

    private bool TryReadPageBox(string key, out PdfPageBox? box) {
        box = null;
        PdfObject? value = GetInheritedValue(key);
        var array = ResolveArray(value);
        if (array is null || array.Items.Count < 4) {
            return false;
        }

        if (ResolveObject(array.Items[0]) is not PdfNumber left ||
            ResolveObject(array.Items[1]) is not PdfNumber bottom ||
            ResolveObject(array.Items[2]) is not PdfNumber right ||
            ResolveObject(array.Items[3]) is not PdfNumber top) {
            return false;
        }

        if (!IsFinite(left.Value) ||
            !IsFinite(bottom.Value) ||
            !IsFinite(right.Value) ||
            !IsFinite(top.Value) ||
            right.Value <= left.Value ||
            top.Value <= bottom.Value) {
            return false;
        }

        box = new PdfPageBox(key, left.Value, bottom.Value, right.Value, top.Value);
        return true;
    }

    private PdfObject? GetDirectValue(string key) {
        return _pageDict.Items.TryGetValue(key, out PdfObject? value) ? value : null;
    }

    private string? TryReadDirectName(string key) {
        return ResolveObject(GetDirectValue(key)) is PdfName name && !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private double? TryReadDirectPositiveNumber(string key) {
        return ResolveObject(GetDirectValue(key)) is PdfNumber number && IsFinite(number.Value) && number.Value > 0
            ? number.Value
            : null;
    }

    private double? TryReadDirectPositiveOrZeroNumber(string key) {
        return ResolveObject(GetDirectValue(key)) is PdfNumber number && IsFinite(number.Value) && number.Value >= 0
            ? number.Value
            : null;
    }

    private bool HasDirectKey(string key) {
        PdfObject? value = GetDirectValue(key);
        return value is not null && ResolveObject(value) is not PdfNull;
    }

    private PdfPageTransition? ReadTransition(PdfObject? value) {
        var dictionary = ResolveDictionary(value);
        if (dictionary is null) {
            return null;
        }

        string? style = TryReadName(dictionary, "S");
        double? durationSeconds = TryReadNonNegativeNumber(dictionary, "D");
        string? dimension = TryReadName(dictionary, "Dm");
        string? motion = TryReadName(dictionary, "M");
        int? direction = TryReadInteger(dictionary, "Di");
        double? scale = TryReadNonNegativeNumber(dictionary, "SS");
        bool? isFlyAreaOpaque = TryReadBoolean(dictionary, "B");

        return new PdfPageTransition(style, durationSeconds, dimension, motion, direction, scale, isFlyAreaOpaque);
    }

    private string? TryReadName(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(value) is PdfName name &&
            !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private double? TryReadNonNegativeNumber(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(value) is PdfNumber number &&
            IsFinite(number.Value) &&
            number.Value >= 0
            ? number.Value
            : null;
    }

    private int? TryReadInteger(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(value) is not PdfNumber number ||
            !IsFinite(number.Value) ||
            Math.Truncate(number.Value) != number.Value ||
            number.Value < int.MinValue ||
            number.Value > int.MaxValue) {
            return null;
        }

        return (int)number.Value;
    }

    private bool? TryReadBoolean(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(value) is PdfBoolean boolean
            ? boolean.Value
            : null;
    }

    private static bool IsFinite(double value) {
        return !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
