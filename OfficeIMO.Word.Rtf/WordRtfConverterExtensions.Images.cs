namespace OfficeIMO.Word.Rtf;

using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;

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
        if (source._paragraph.ChildElements.Any(child =>
                child is not DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties &&
                child is not DocumentFormat.OpenXml.Wordprocessing.Run)) {
            return false;
        }

        List<DocumentFormat.OpenXml.Wordprocessing.Run> runs = source._paragraph
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Run>()
            .ToList();
        if (runs.Any(run => !string.IsNullOrEmpty(
                new WordParagraph(source._document, source._paragraph, run).Text))) {
            return false;
        }

        bool copied = false;
        foreach (DocumentFormat.OpenXml.Wordprocessing.Run sourceRun in runs) {
            var run = new WordParagraph(source._document, source._paragraph, sourceRun);
            foreach (WordImage wordImage in run.EnumerateImages()) {
                RtfImage? image = CreateRtfImage(wordImage, out _, out _);
                if (image == null) continue;
                CopyImage(image, addImage(image));
                copied = true;
            }
        }
        return copied;
    }

    private static RtfImage? CreateRtfImage(
        WordImage source,
        out OfficeImageFormat sourceFormat,
        out bool animationDiscarded) {
        sourceFormat = OfficeImageFormat.Unknown;
        animationDiscarded = false;
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
                out sourceFormat,
                out animationDiscarded)) {
            return null;
        }

        double visibleRatioX = 1d - (source.CropLeft ?? 0) / 100000d - (source.CropRight ?? 0) / 100000d;
        double visibleRatioY = 1d - (source.CropTop ?? 0) / 100000d - (source.CropBottom ?? 0) / 100000d;
        int? goalWidth = ToTwips(source.Width.HasValue && visibleRatioX > 0 ? source.Width / visibleRatioX : source.Width);
        int? goalHeight = ToTwips(source.Height.HasValue && visibleRatioY > 0 ? source.Height / visibleRatioY : source.Height);
        OfficeImageReader.TryValidateContent(payload, null, out OfficeImageInfo? info);
        var image = new RtfImage(format, payload) {
            SourceWidth = info?.Width > 0 ? info.Width : null,
            SourceHeight = info?.Height > 0 ? info.Height : null,
            DesiredWidthTwips = goalWidth,
            DesiredHeightTwips = goalHeight,
            CropLeftTwips = ToCropTwips(source.CropLeft, goalWidth),
            CropRightTwips = ToCropTwips(source.CropRight, goalWidth),
            CropTopTwips = ToCropTwips(source.CropTop, goalHeight),
            CropBottomTwips = ToCropTwips(source.CropBottom, goalHeight),
            Description = source.Description
        };
        return image;
    }

    private static void CopyImage(RtfImage source, RtfImage destination) {
        destination.SourceWidth = source.SourceWidth;
        destination.SourceHeight = source.SourceHeight;
        destination.DesiredWidthTwips = source.DesiredWidthTwips;
        destination.DesiredHeightTwips = source.DesiredHeightTwips;
        destination.ScaleXPercent = source.ScaleXPercent;
        destination.ScaleYPercent = source.ScaleYPercent;
        destination.CropLeftTwips = source.CropLeftTwips;
        destination.CropRightTwips = source.CropRightTwips;
        destination.CropTopTwips = source.CropTopTwips;
        destination.CropBottomTwips = source.CropBottomTwips;
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
        if (!TryGetWordImageLayout(image, payload, fileName, out RtfImageLayout? layout)) return;
        WordImage output = paragraph.InsertImage(
            stream,
            fileName,
            layout!.VisibleWidthTwips * PixelsPerTwip,
            layout.VisibleHeightTwips * PixelsPerTwip,
            WordImageTextWrapping.InLineWithText,
            image.Description ?? string.Empty);
        output.CropLeft = ToWordCrop(image.CropLeftTwips, layout.WidthTwips);
        output.CropRight = ToWordCrop(image.CropRightTwips, layout.WidthTwips);
        output.CropTop = ToWordCrop(image.CropTopTwips, layout.HeightTwips);
        output.CropBottom = ToWordCrop(image.CropBottomTwips, layout.HeightTwips);
    }

    private static int? ToCropTwips(int? fraction, int? goal) => fraction.HasValue && goal.HasValue ? checked((int)Math.Round(fraction.Value / 100000d * goal.Value, MidpointRounding.AwayFromZero)) : null;
    private static int? ToWordCrop(int? crop, double? goal) => crop.HasValue && goal > 0 ? checked((int)Math.Round(crop.Value / goal.Value * 100000d, MidpointRounding.AwayFromZero)) : null;

    private static bool CanWriteToWord(RtfImage image) =>
        TryGetWordImagePayload(image, out byte[] payload, out string fileName) && TryGetWordImageLayout(image, payload, fileName, out _);

    private static bool TryGetWordImageLayout(RtfImage image, byte[] payload, string fileName, out RtfImageLayout? layout) {
        OfficeImageReader.TryValidateContent(payload, fileName, out OfficeImageInfo? info);
        try {
            layout = image.ResolveLayout(info?.Width > 0 ? info.Width * 15d : null, info?.Height > 0 ? info.Height * 15d : null);
            return true;
        } catch (InvalidDataException) {
            layout = null;
            return false;
        }
    }

    private static bool TryCreateRtfImagePayload(
        byte[] bytes,
        string? fileName,
        out RtfImageFormat format,
        out byte[] payload,
        out OfficeImageFormat sourceFormat,
        out bool animationDiscarded) {
        format = RtfImageFormat.Unknown;
        payload = Array.Empty<byte>();
        sourceFormat = OfficeImageFormat.Unknown;
        animationDiscarded = false;
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
                    if (OfficeImagePngConverter.TryConvertToPng(
                            bytes,
                            options: null,
                            out byte[] normalized,
                            out OfficeRasterDecodeInfo decodeInfo)) {
                        format = RtfImageFormat.Png;
                        payload = normalized;
                        animationDiscarded = decodeInfo.AnimationDiscarded;
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

    private static int? ToTwips(double? pixels) {
        if (!pixels.HasValue) return null;
        return (int)Math.Round(pixels.Value * TwipsPerPixel, MidpointRounding.AwayFromZero);
    }

}
