using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static bool TryGetJpxPayload(PdfStream stream, Dictionary<int, PdfIndirectObject> objects,
        string colorSpace, bool hasUnsupportedColorSpaceDeclaration, int maximumBytes, out byte[] payload, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        payload = Array.Empty<byte>();
        // The unchanged JPEG 2000 file is kept for opaque Gray/RGB images whose PDF dictionary adds nothing the
        // codestream lacks. PDF masks, alternate colour spaces and output intents go through sample normalization.
        if (hasUnsupportedColorSpaceDeclaration ||
            colorSpace is not ("" or "DeviceGray" or "G" or "DeviceRGB" or "RGB") ||
            GetTransparencyMaskKind(stream.Dictionary, objects) != null ||
            PdfImageMaskNormalizer.IsImageMask(stream, objects)) return false;
        if (stream.Dictionary.Items.TryGetValue("SMaskInData", out PdfObject? embeddedMask) &&
            PdfObjectLookup.ResolveChain(objects, embeddedMask) is not (PdfNull or PdfNumber { Value: 0 })) return false;
        if (!TryDecodeJpxPrefix(stream, objects, maximumBytes, out payload, cancellationToken)) return false;
        if (!OfficeJpeg2000Header.TryValidateOpaquePayload(payload, cancellationToken, out int components, out int width, out int height) ||
            stream.Dictionary.Get<PdfNumber>("Width")?.Value != width ||
            stream.Dictionary.Get<PdfNumber>("Height")?.Value != height || (long)width * height * 4L > maximumBytes) return false;
        return colorSpace == "" || (components == 1 ? colorSpace is "DeviceGray" or "G" : colorSpace is "DeviceRGB" or "RGB");
    }

    private static bool IsJpxFilterChain(PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, dictionary.Items.TryGetValue("Filter", out PdfObject? value) ? value : null);
        if (resolved is PdfName name) return name.Name == "JPXDecode";
        return resolved is PdfArray array && array.Items.Count > 0 &&
            PdfObjectLookup.ResolveChain(objects, array.Items[array.Items.Count - 1]) is PdfName { Name: "JPXDecode" };
    }

    // Applies the filters before JPXDecode (ISO 32000-1 7.4.9: JPXDecode is the last filter) and returns the JPEG 2000 data.
    private static bool TryDecodeJpxPrefix(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, int maximumBytes,
        out byte[] payload, CancellationToken cancellationToken) {
        payload = Array.Empty<byte>();
        PdfObject? resolved = PdfObjectLookup.ResolveChain(objects, stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? value) ? value : null);
        var filters = new List<PdfObject>();
        if (resolved is PdfName name) filters.Add(name);
        else if (resolved is PdfArray array) filters.AddRange(array.Items);
        else return false;
        if (filters.Count == 0 || PdfObjectLookup.ResolveChain(objects, filters[filters.Count - 1]) is not PdfName { Name: "JPXDecode" }) return false;
        if (filters.Count == 1) {
            payload = stream.Data;
            return true;
        }
        var prefix = new PdfDictionary();
        var prefixFilters = new PdfArray();
        for (int i = 0; i < filters.Count - 1; i++) prefixFilters.Items.Add(filters[i]);
        prefix.Items["Filter"] = prefixFilters;
        if (stream.Dictionary.Items.TryGetValue("DecodeParms", out PdfObject? parameters)) {
            PdfObject? resolvedParameters = PdfObjectLookup.ResolveChain(objects, parameters);
            if (resolvedParameters is PdfArray parameterArray) {
                if (parameterArray.Items.Count != filters.Count ||
                    PdfObjectLookup.ResolveChain(objects, parameterArray.Items[filters.Count - 1]) is not PdfNull) return false;
                var prefixParameters = new PdfArray();
                for (int i = 0; i < filters.Count - 1; i++) prefixParameters.Items.Add(parameterArray.Items[i]);
                prefix.Items["DecodeParms"] = prefixParameters;
            } else if (resolvedParameters is not PdfNull) return false;
        }
        try {
            payload = Filters.StreamDecoder.DecodeRequired(prefix, stream.Data, objects, maximumBytes, cancellationToken);
            return true;
        } catch (InvalidDataException) {
            return false;
        }
    }

    /// <summary>
    /// Decodes a JPXDecode image with the managed JPEG 2000 decoder into an equivalent unfiltered 8-bit image stream.
    /// A /ColorSpace entry overrides the JPX colour specification; without one, the JP2 colour box (enumerated or
    /// ICC) selects DeviceGray, DeviceRGB, DeviceCMYK or ICCBased. /Decode is ignored and /SMaskInData turns the
    /// opacity channel into a soft mask, un-premultiplied for value 2 (ISO 32000-1 8.9.5 and Table 89).
    /// </summary>
    private static bool TryNormalizeJpxImage(PdfStream stream, Dictionary<int, PdfIndirectObject> objects,
        PdfObject? declaredColorSpace, int declaredColorCount, bool indexed, int maximumBytes,
        out PdfStream normalized, out PdfObject? colorSpaceObject, CancellationToken cancellationToken) {
        normalized = stream;
        colorSpaceObject = declaredColorSpace;
        if (!TryDecodeJpxPrefix(stream, objects, maximumBytes, out byte[] payload, cancellationToken)) return false;
        var options = new OfficeJpeg2000DecodeOptions {
            MaximumPixels = OfficeRasterGuards.MaximumPixels,
            MaximumWorkingBytes = maximumBytes,
            ScaleToEightBits = !indexed,
            ApplyContainerChannels = !indexed,
            CancellationToken = cancellationToken
        };
        if (!OfficeJpeg2000Decoder.TryDecode(payload, options, out OfficeJpeg2000Image? image) || image == null) return false;

        int colorCount = declaredColorCount;
        if (declaredColorSpace == null) {
            colorCount = image.ColorChannels;
            switch (image.ColorSpace) {
                case OfficeJpeg2000ColorSpace.Icc when image.IccProfile != null && colorCount is 1 or 3 or 4: {
                    var profileDictionary = new PdfDictionary();
                    profileDictionary.Items["N"] = new PdfNumber(colorCount);
                    var iccBased = new PdfArray();
                    iccBased.Items.Add(new PdfName("ICCBased"));
                    iccBased.Items.Add(new PdfStream(profileDictionary, image.IccProfile));
                    colorSpaceObject = iccBased;
                    break;
                }
                default:
                    colorSpaceObject = colorCount switch {
                        1 => new PdfName("DeviceGray"),
                        3 => new PdfName("DeviceRGB"),
                        4 => new PdfName("DeviceCMYK"),
                        _ => null
                    };
                    break;
            }
            if (colorSpaceObject == null) return false;
        }
        // An /Indexed image must be a single-component codestream (ISO 32000-1 Table 89); other declared spaces
        // take the leading colour channels.
        if (colorCount < 1 || colorCount > image.ColorChannels || indexed && image.ColorChannels != 1) return false;

        int width = image.Width, height = image.Height;
        if (stream.Dictionary.Get<PdfNumber>("Width")?.Value is double declaredWidth && declaredWidth != width ||
            stream.Dictionary.Get<PdfNumber>("Height")?.Value is double declaredHeight && declaredHeight != height) return false;
        long pixels = (long)width * height;
        if (pixels * colorCount > maximumBytes) return false;
        var samples = new byte[pixels * colorCount];
        byte[]? alpha = null;
        int smaskInData = stream.Dictionary.Items.TryGetValue("SMaskInData", out PdfObject? smaskInDataObject) &&
            PdfObjectLookup.ResolveChain(objects, smaskInDataObject) is PdfNumber number ? (int)number.Value : 0;
        bool hasSoftMask = PdfImageMaskSemantics.HasSoftMask(stream.Dictionary, objects);
        if (smaskInData != 0 && image.HasAlpha && !hasSoftMask) alpha = new byte[pixels];
        int channels = image.Channels;
        for (long pixel = 0; pixel < pixels; pixel++) {
            if ((pixel & 0xFFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            long source = pixel * channels, target = pixel * colorCount;
            byte opacity = image.HasAlpha ? image.Samples[source + image.ColorChannels] : (byte)255;
            for (int component = 0; component < colorCount; component++) {
                int sample = image.Samples[source + component];
                if (alpha != null && smaskInData == 2 && opacity != 255)
                    sample = opacity == 0 ? 0 : Math.Min(255, (sample * 255 + opacity / 2) / opacity);
                samples[target + component] = (byte)sample;
            }
            if (alpha != null) alpha[pixel] = opacity;
        }

        var dictionary = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> pair in stream.Dictionary.Items) dictionary.Items.Add(pair.Key, pair.Value);
        foreach (string key in new[] { "Filter", "DecodeParms", "F", "DP", "SMaskInData", "Decode", "Length" }) dictionary.Items.Remove(key);
        dictionary.Items["Width"] = new PdfNumber(width);
        dictionary.Items["Height"] = new PdfNumber(height);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        dictionary.Items["ColorSpace"] = colorSpaceObject!;
        if (dictionary.Items.TryGetValue("Mask", out PdfObject? mask) && PdfObjectLookup.ResolveChain(objects, mask) is PdfArray &&
            Array.Exists(image.Precisions, static precision => precision != 8)) {
            // Colour-key ranges are expressed in the codestream precision; only eight-bit samples keep them exact.
            return false;
        }
        if (alpha != null) {
            var maskDictionary = new PdfDictionary();
            maskDictionary.Items["Type"] = new PdfName("XObject");
            maskDictionary.Items["Subtype"] = new PdfName("Image");
            maskDictionary.Items["Width"] = new PdfNumber(width);
            maskDictionary.Items["Height"] = new PdfNumber(height);
            maskDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
            maskDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            dictionary.Items["SMask"] = new PdfStream(maskDictionary, alpha);
        }
        normalized = new PdfStream(dictionary, samples);
        return true;
    }

    private static bool TryBuildJpxPngFile(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, PdfObject? declaredColorSpace,
        string colorSpace, int maxDecodedStreamBytes, OfficeIccRenderingIntent renderingIntent,
        PdfOutputIntentColorTransform? outputIntentColorTransform, Func<int, long, bool>? colorFunctionEvaluationBudget,
        PdfColorFunctionResolutionContext? functionResolutionContext, out byte[] pngBytes, out string? transparencyMaskKind,
        out bool transparencyMaskResolved, CancellationToken cancellationToken) {
        pngBytes = Array.Empty<byte>();
        transparencyMaskKind = null;
        transparencyMaskResolved = false;
        bool indexed = declaredColorSpace != null && PdfIndexedImageNormalizer.IsIndexedColorSpace(declaredColorSpace, objects);
        int declaredCount = 0;
        if (declaredColorSpace != null && !indexed) {
            if (!PdfImageColorSpaceNormalization.TryResolve(declaredColorSpace, colorSpace, objects, maxDecodedStreamBytes, renderingIntent,
                    outputIntentColorTransform, colorFunctionEvaluationBudget, functionResolutionContext, out PdfImageColorSpaceNormalization declared)) return false;
            declaredCount = declared.SourceColorCount;
        } else if (indexed) {
            declaredCount = 1;
        }
        if (!TryNormalizeJpxImage(stream, objects, declaredColorSpace, declaredCount, indexed, maxDecodedStreamBytes,
                out PdfStream normalized, out PdfObject? normalizedColorSpace, cancellationToken)) return false;
        string normalizedName = declaredColorSpace != null ? colorSpace : GetNameOrEmpty(normalizedColorSpace, objects);
        int width = (int)normalized.Dictionary.Get<PdfNumber>("Width")!.Value;
        int height = (int)normalized.Dictionary.Get<PdfNumber>("Height")!.Value;
        if (!TryBuildPngFile(normalized, width, height, 8, normalizedColorSpace, normalizedName, string.Empty, objects, maxDecodedStreamBytes,
                renderingIntent, outputIntentColorTransform, colorFunctionEvaluationBudget, functionResolutionContext, out pngBytes,
                cancellationToken: cancellationToken)) return false;
        transparencyMaskKind = GetTransparencyMaskKind(normalized.Dictionary, objects);
        transparencyMaskResolved = IsTransparencyMaskResolvedByPngNormalization(transparencyMaskKind, normalized.Dictionary,
            normalizedColorSpace, normalizedName, 8, objects, maxDecodedStreamBytes);
        return true;
    }
}
