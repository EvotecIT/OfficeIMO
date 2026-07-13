namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private const double PixelsPerTwip = 96D / 1440D;
    private const double TwipsPerPixel = 1440D / 96D;

    private static bool TryCopyImageBlock(WordParagraph source, RtfDocument destination) {
        if (!string.IsNullOrEmpty(source.Text)) return false;

        RtfImage? image = CreateRtfImage(source);
        if (image == null) return false;

        CopyImage(image, destination.AddImage(image.Format, image.Data));
        return true;
    }

    private static bool TryCopyImageBlock(WordParagraph source, RtfSection destination) {
        if (!string.IsNullOrEmpty(source.Text)) return false;

        RtfImage? image = CreateRtfImage(source);
        if (image == null) return false;

        CopyImage(image, destination.AddImage(image.Format, image.Data));
        return true;
    }

    private static RtfImage? CreateRtfImage(WordParagraph source) {
        if (!source.IsImage || source.Image == null || source.Image.IsExternal) {
            return null;
        }

        byte[] bytes;
        try {
            bytes = source.Image.ToBytes();
        } catch (InvalidOperationException) {
            return null;
        }

        if (bytes.Length == 0) {
            return null;
        }

        var image = new RtfImage(DetectRtfImageFormat(bytes, source.Image.FileName), bytes) {
            SourceWidth = ToNullableInt(source.Image.Width),
            SourceHeight = ToNullableInt(source.Image.Height),
            DesiredWidthTwips = ToTwips(source.Image.Width),
            DesiredHeightTwips = ToTwips(source.Image.Height),
            Description = source.Image.Description
        };
        return image;
    }

    private static void CopyImage(RtfImage source, RtfImage destination) {
        destination.SourceWidth = source.SourceWidth;
        destination.SourceHeight = source.SourceHeight;
        destination.DesiredWidthTwips = source.DesiredWidthTwips;
        destination.DesiredHeightTwips = source.DesiredHeightTwips;
        destination.Description = source.Description;
    }

    private static void AppendImage(WordDocument document, RtfImage image) {
        WordParagraph paragraph = document.AddParagraph();
        AppendImage(paragraph, image);
    }

    private static void AppendImage(WordSection section, RtfImage image) {
        WordParagraph paragraph = section.AddParagraph(newRun: true);
        AppendImage(paragraph, image);
    }

    private static void AppendImage(WordParagraph paragraph, RtfImage image) {
        if (!CanWriteToWord(image)) {
            return;
        }

        using var stream = new MemoryStream(image.Data);
        paragraph.AddImage(
            stream,
            GetImageFileName(image.Format),
            ToPixels(image.DesiredWidthTwips),
            ToPixels(image.DesiredHeightTwips),
            WrapTextImage.InLineWithText,
            image.Description ?? string.Empty);
    }

    private static bool CanWriteToWord(RtfImage image) {
        return image.Data.Length > 0 &&
            (image.Format == RtfImageFormat.Png ||
             image.Format == RtfImageFormat.Jpeg ||
             image.Format == RtfImageFormat.Dib ||
             image.Format == RtfImageFormat.Wmf ||
             image.Format == RtfImageFormat.Emf);
    }

    private static RtfImageFormat DetectRtfImageFormat(byte[] bytes, string? fileName) {
        if (bytes.Length >= 8 &&
            bytes[0] == 0x89 &&
            bytes[1] == 0x50 &&
            bytes[2] == 0x4E &&
            bytes[3] == 0x47 &&
            bytes[4] == 0x0D &&
            bytes[5] == 0x0A &&
            bytes[6] == 0x1A &&
            bytes[7] == 0x0A) {
            return RtfImageFormat.Png;
        }

        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF) {
            return RtfImageFormat.Jpeg;
        }

        string extension = Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant();
        switch (extension) {
            case ".png":
                return RtfImageFormat.Png;
            case ".jpg":
            case ".jpeg":
                return RtfImageFormat.Jpeg;
            case ".bmp":
            case ".dib":
                return RtfImageFormat.Dib;
            case ".wmf":
                return RtfImageFormat.Wmf;
            case ".emf":
                return RtfImageFormat.Emf;
            default:
                return RtfImageFormat.Unknown;
        }
    }

    private static string GetImageFileName(RtfImageFormat format) {
        string extension = format switch {
            RtfImageFormat.Jpeg => "jpg",
            RtfImageFormat.Dib => "bmp",
            RtfImageFormat.Wmf => "wmf",
            RtfImageFormat.Emf => "emf",
            _ => "png"
        };
        return "rtf-image." + extension;
    }

    private static int? ToNullableInt(double? value) {
        if (!value.HasValue) return null;
        return (int)Math.Round(value.Value, MidpointRounding.AwayFromZero);
    }

    private static int? ToTwips(double? pixels) {
        if (!pixels.HasValue) return null;
        return (int)Math.Round(pixels.Value * TwipsPerPixel, MidpointRounding.AwayFromZero);
    }

    private static double? ToPixels(int? twips) {
        if (!twips.HasValue) return null;
        return Math.Round(twips.Value * PixelsPerTwip, 2, MidpointRounding.AwayFromZero);
    }
}
