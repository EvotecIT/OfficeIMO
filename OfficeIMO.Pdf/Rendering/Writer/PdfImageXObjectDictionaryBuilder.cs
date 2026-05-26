using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfImageXObjectDictionaryBuilder {
    internal static string BuildStreamDictionary(PdfWriter.PdfImageStream image, int? softMaskObjectId = null) {
        ValidateImageStream(image);
        string softMask = softMaskObjectId.HasValue
            ? " /SMask " + PdfSyntaxEscaper.IndirectReference(softMaskObjectId.Value)
            : string.Empty;

        return "<< /Type /XObject /Subtype /Image /Width " +
               image.PixelWidth.ToString(CultureInfo.InvariantCulture) +
               " /Height " +
               image.PixelHeight.ToString(CultureInfo.InvariantCulture) +
               image.DictionarySuffix +
               softMask +
               " /Length " +
               image.Data.Length.ToString(CultureInfo.InvariantCulture) +
               " >>";
    }

    internal static PdfStream BuildStreamObject(PdfWriter.PdfImageStream image, int? softMaskObjectNumber = null) {
        ValidateImageStream(image);

        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Image");
        dictionary.Items["Width"] = new PdfNumber(image.PixelWidth);
        dictionary.Items["Height"] = new PdfNumber(image.PixelHeight);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);

        if (image.DictionarySuffix.Contains("/DeviceGray")) {
            dictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
        } else {
            dictionary.Items["ColorSpace"] = new PdfName("DeviceRGB");
        }

        if (image.DictionarySuffix.Contains("/FlateDecode")) {
            dictionary.Items["Filter"] = new PdfName("FlateDecode");
            var decodeParms = new PdfDictionary();
            decodeParms.Items["Predictor"] = new PdfNumber(15);
            decodeParms.Items["Colors"] = new PdfNumber(image.DictionarySuffix.Contains("/DeviceGray") ? 1 : 3);
            decodeParms.Items["BitsPerComponent"] = new PdfNumber(8);
            decodeParms.Items["Columns"] = new PdfNumber(image.PixelWidth);
            dictionary.Items["DecodeParms"] = decodeParms;
        } else {
            dictionary.Items["Filter"] = new PdfName("DCTDecode");
        }

        if (softMaskObjectNumber.HasValue) {
            dictionary.Items["SMask"] = new PdfReference(softMaskObjectNumber.Value, 0);
        }

        return new PdfStream(dictionary, image.Data);
    }

    private static void ValidateImageStream(PdfWriter.PdfImageStream image) {
        Guard.NotNull(image, nameof(image));
        Guard.NotNullOrEmpty(image.Data, nameof(image.Data));
        Guard.NotNullOrWhiteSpace(image.DictionarySuffix, nameof(image.DictionarySuffix));
        if (image.PixelWidth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(image), image.PixelWidth, "PDF image width must be positive.");
        }

        if (image.PixelHeight <= 0) {
            throw new ArgumentOutOfRangeException(nameof(image), image.PixelHeight, "PDF image height must be positive.");
        }
    }
}
