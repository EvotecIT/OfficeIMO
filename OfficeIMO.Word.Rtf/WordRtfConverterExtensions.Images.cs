namespace OfficeIMO.Word.Rtf;

using OfficeIMO.Drawing;

public static partial class WordRtfConverterExtensions {
    private const double PixelsPerTwip = 96D / 1440D;
    private const double TwipsPerPixel = 1440D / 96D;

    private static bool TryCopyImageBlock(WordParagraph source, RtfDocument destination) {
        if (!string.IsNullOrEmpty(source.Text)) return false;
        return TryCopyImageBlocks(source, image => destination.AddImage(image.Format, image.Data));
    }

    private static bool TryCopyImageBlock(WordParagraph source, RtfSection destination) {
        if (!string.IsNullOrEmpty(source.Text)) return false;
        return TryCopyImageBlocks(source, image => destination.AddImage(image.Format, image.Data));
    }

    private static bool TryCopyImageBlocks(WordParagraph source, Func<RtfImage, RtfImage> addImage) {
        bool copied = false;
        foreach (WordImage wordImage in source.EnumerateImages()) {
            RtfImage? image = CreateRtfImage(wordImage, out _);
            if (image == null) continue;
            CopyImage(image, addImage(image));
            copied = true;
        }
        return copied;
    }

    private static RtfImage? CreateRtfImage(WordImage source, out OfficeImageFormat sourceFormat) {
        sourceFormat = OfficeImageFormat.Unknown;
        if (source.IsExternal) {
            return null;
        }

        byte[] bytes;
        try {
            bytes = source.ToBytes();
        } catch (InvalidOperationException) {
            return null;
        }

        if (bytes.Length == 0) {
            return null;
        }

        if (!TryCreateRtfImagePayload(
                bytes,
                source.FileName,
                out RtfImageFormat format,
                out byte[] payload,
                out sourceFormat)) {
            return null;
        }

        var image = new RtfImage(format, payload) {
            SourceWidth = ToNullableInt(source.Width),
            SourceHeight = ToNullableInt(source.Height),
            DesiredWidthTwips = ToTwips(source.Width),
            DesiredHeightTwips = ToTwips(source.Height),
            Description = source.Description
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
        if (!TryGetWordImagePayload(image, out byte[] payload, out string fileName)) {
            return;
        }

        using var stream = new MemoryStream(payload);
        paragraph.AddImage(
            stream,
            fileName,
            ToPixels(image.DesiredWidthTwips),
            ToPixels(image.DesiredHeightTwips),
            WordImageTextWrapping.InLineWithText,
            image.Description ?? string.Empty);
    }

    private static bool CanWriteToWord(RtfImage image) =>
        TryGetWordImagePayload(image, out _, out _);

    private static bool TryCreateRtfImagePayload(
        byte[] bytes,
        string? fileName,
        out RtfImageFormat format,
        out byte[] payload,
        out OfficeImageFormat sourceFormat) {
        format = RtfImageFormat.Unknown;
        payload = Array.Empty<byte>();
        sourceFormat = OfficeImageFormat.Unknown;
        if (OfficeImageReader.TryValidateContent(bytes, fileName, out OfficeImageInfo info)) {
            sourceFormat = info.Format;
            switch (info.Format) {
                case OfficeImageFormat.Png:
                    format = RtfImageFormat.Png;
                    payload = bytes;
                    return true;
                case OfficeImageFormat.Jpeg:
                    format = RtfImageFormat.Jpeg;
                    payload = bytes;
                    return true;
                case OfficeImageFormat.Wmf:
                    format = RtfImageFormat.Wmf;
                    payload = bytes;
                    return true;
                case OfficeImageFormat.Emf:
                    format = RtfImageFormat.Emf;
                    payload = bytes;
                    return true;
                default:
                    if (OfficeImagePngConverter.TryConvertToPng(bytes, out byte[] normalized)) {
                        format = RtfImageFormat.Png;
                        payload = normalized;
                        return true;
                    }
                    break;
            }
        }

        if (string.Equals(Path.GetExtension(fileName), ".dib", StringComparison.OrdinalIgnoreCase) &&
            OfficeImagePngConverter.TryConvertDibToPng(bytes, out byte[] dibPng)) {
            format = RtfImageFormat.Png;
            payload = dibPng;
            return true;
        }
        return false;
    }

    private static bool TryGetWordImagePayload(RtfImage image, out byte[] payload, out string fileName) {
        payload = Array.Empty<byte>();
        fileName = string.Empty;
        if (image.Data.Length == 0) return false;

        if (image.Format == RtfImageFormat.Dib) {
            if (!OfficeImagePngConverter.TryConvertDibToPng(image.Data, out payload)) return false;
            fileName = "rtf-image.png";
            return true;
        }

        OfficeImageFormat expected = image.Format switch {
            RtfImageFormat.Png => OfficeImageFormat.Png,
            RtfImageFormat.Jpeg => OfficeImageFormat.Jpeg,
            RtfImageFormat.Wmf => OfficeImageFormat.Wmf,
            RtfImageFormat.Emf => OfficeImageFormat.Emf,
            _ => OfficeImageFormat.Unknown
        };
        if (expected == OfficeImageFormat.Unknown ||
            !OfficeImageReader.TryValidateContent(image.Data, OfficeImageInfo.GetDefaultExtension(expected), out OfficeImageInfo info) ||
            info.Format != expected) {
            return false;
        }

        payload = image.Data;
        fileName = "rtf-image" + OfficeImageInfo.GetDefaultExtension(expected);
        return true;
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
