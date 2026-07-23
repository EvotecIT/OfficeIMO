using OfficeIMO.Pdf.Filters;
using OfficeIMO.Drawing;
using System.Globalization;
using System.IO.Compression;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const double ImagePixelRewriteTransformTolerance = 0.0000001D;

    private static bool RewriteMatchedImagePixels(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        PdfObject contentsObject,
        ImageRedactionTarget[] targets,
        PdfRedactionApplyOptions options,
        IReadOnlyDictionary<int, int> referenceCounts,
        List<PdfRedactionMatch> removedMatches,
        ref int nextObjectNumber) {
        if (targets.Length == 0) {
            return false;
        }

        PdfDictionary? resources = GetInheritedDictionary(objects, pageDictionary, "Resources");
        if (resources is null || !resources.Items.ContainsKey("XObject")) {
            return false;
        }

        PdfDictionary xObjects = PdfPageResourceHelper.EnsurePageXObjects(objects, pageDictionary, "redaction image pixel rewrite");
        resources = ResolveDictionary(objects, pageDictionary.Items.TryGetValue("Resources", out PdfObject? pageResources) ? pageResources : null) ?? resources;
        bool changed = false;
        PdfObject currentContentsObject = contentsObject;
        foreach (PdfReference reference in EnumerateContentReferences(objects, contentsObject)) {
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) ||
                indirect.Value is not PdfStream stream ||
                stream.DecodingFailed) {
                continue;
            }

            string content = PdfEncoding.Latin1GetString(StreamDecoder.Decode(stream.Dictionary, stream.Data, objects));
            ImagePixelRewriteContentResult result = RewriteImagePixelsInContent(objects, resources, xObjects, content, targets, options, Matrix2D.Identity, referenceCounts, new HashSet<int>(), removedMatches, ref nextObjectNumber);
            if (!string.Equals(result.Content, content, StringComparison.Ordinal)) {
                PdfReference targetReference = reference;
                if (IsSharedReference(referenceCounts, reference)) {
                    targetReference = CloneIndirectObject(objects, reference, indirect, ref nextObjectNumber);
                    ReplacePageContentReference(objects, pageDictionary, currentContentsObject, reference, targetReference);
                    currentContentsObject = pageDictionary.Items.TryGetValue("Contents", out PdfObject? updatedContentsObject)
                        ? updatedContentsObject
                        : currentContentsObject;
                }

                objects[targetReference.ObjectNumber] = new PdfIndirectObject(targetReference.ObjectNumber, targetReference.Generation, new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(result.Content)));
            }

            changed = result.HasChanges || changed;
        }

        return changed;
    }

    private static ImagePixelRewriteContentResult RewriteImagePixelsInContent(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        PdfDictionary xObjects,
        string content,
        ImageRedactionTarget[] targets,
        PdfRedactionApplyOptions options,
        Matrix2D baseTransform,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        List<PdfRedactionMatch> removedMatches,
        ref int nextObjectNumber) {
        bool changed = false;
        string rewrittenContent = content;
        ImageResourceInvocation[] invocations = ExtractImageResourceInvocations(content);
        for (int invocationIndex = invocations.Length - 1; invocationIndex >= 0; invocationIndex--) {
            ImageResourceInvocation invocation = invocations[invocationIndex];
            Matrix2D invocationTransform = Matrix2D.Multiply(baseTransform, invocation.Transform);
            if (TryGetImageXObject(objects, xObjects, invocation.Name, out PdfReference imageReference, out PdfStream imageStream)) {
                if (!TryFindImageTarget(invocation.Name, invocationTransform, targets, out ImageRedactionTarget target) ||
                    !CanRewriteImagePlacementPixels(invocationTransform)) {
                    continue;
                }

                bool repeatedInvocation = CountResourceInvocations(content, invocation.Name) != 1;
                if (repeatedInvocation &&
                    PdfObjectLookup.TryGet(objects, imageReference, out PdfIndirectObject? repeatedSourceIndirect)) {
                    string resourceName = CreateUniqueResourceName(xObjects, invocation.Name);
                    imageReference = CloneIndirectObject(objects, imageReference, repeatedSourceIndirect, ref nextObjectNumber);
                    xObjects.Items[resourceName] = imageReference;
                    imageStream = (PdfStream)objects[imageReference.ObjectNumber].Value;
                    rewrittenContent = ReplaceInvocationResourceName(rewrittenContent, invocation, resourceName);
                    changed = true;
                } else if (IsCurrentlySharedReference(objects, referenceCounts, imageReference) &&
                    PdfObjectLookup.TryGet(objects, imageReference, out PdfIndirectObject? sourceIndirect)) {
                    imageReference = CloneIndirectObject(objects, imageReference, sourceIndirect, ref nextObjectNumber);
                    xObjects.Items[invocation.Name] = imageReference;
                    imageStream = (PdfStream)objects[imageReference.ObjectNumber].Value;
                    changed = true;
                }

                if (TryRewriteImageStreamPixels(objects, resources, invocation.Name, imageReference, imageStream, target, invocationTransform, options, ref nextObjectNumber)) {
                    AddRemovedImageTargets(new[] { target }, removedMatches, null);
                    changed = true;
                }

                continue;
            }

            if (!TryGetFormXObject(objects, xObjects, invocation.Name, out PdfReference formReference, out PdfStream formStream) ||
                formStream.DecodingFailed ||
                !activeForms.Add(formReference.ObjectNumber)) {
                continue;
            }

            int activeObjectNumber = formReference.ObjectNumber;
            try {
                bool repeatedInvocation = CountResourceInvocations(content, invocation.Name) != 1;
                if (repeatedInvocation &&
                    PdfObjectLookup.TryGet(objects, formReference, out PdfIndirectObject? repeatedSourceIndirect)) {
                    string resourceName = CreateUniqueResourceName(xObjects, invocation.Name);
                    formReference = CloneIndirectObject(objects, formReference, repeatedSourceIndirect, ref nextObjectNumber);
                    xObjects.Items[resourceName] = formReference;
                    formStream = (PdfStream)objects[formReference.ObjectNumber].Value;

                    ImagePixelRewriteContentResult repeatedResult = RewriteImagePixelsInForm(objects, formReference, formStream, targets, options, invocationTransform, referenceCounts, activeForms, removedMatches, ref nextObjectNumber);
                    if (repeatedResult.HasChanges) {
                        rewrittenContent = ReplaceInvocationResourceName(rewrittenContent, invocation, resourceName);
                        changed = true;
                    }

                    continue;
                }

                if (IsCurrentlySharedReference(objects, referenceCounts, formReference) &&
                    PdfObjectLookup.TryGet(objects, formReference, out PdfIndirectObject? sourceIndirect)) {
                    formReference = CloneIndirectObject(objects, formReference, sourceIndirect, ref nextObjectNumber);
                    xObjects.Items[invocation.Name] = formReference;
                    formStream = (PdfStream)objects[formReference.ObjectNumber].Value;
                    changed = true;
                }

                changed = RewriteImagePixelsInForm(objects, formReference, formStream, targets, options, invocationTransform, referenceCounts, activeForms, removedMatches, ref nextObjectNumber).HasChanges || changed;
            } finally {
                activeForms.Remove(activeObjectNumber);
            }
        }

        return new ImagePixelRewriteContentResult(changed, rewrittenContent);
    }

    private static ImagePixelRewriteContentResult RewriteImagePixelsInForm(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReference formReference,
        PdfStream formStream,
        ImageRedactionTarget[] targets,
        PdfRedactionApplyOptions options,
        Matrix2D invocationTransform,
        IReadOnlyDictionary<int, int> referenceCounts,
        HashSet<int> activeForms,
        List<PdfRedactionMatch> removedMatches,
        ref int nextObjectNumber) {
        PdfDictionary formResources = EnsureFormResources(objects, formStream);
        PdfDictionary formXObjects = EnsureResourceXObjects(objects, formResources);
        Matrix2D formTransform = ApplyFormMatrix(invocationTransform, formStream.Dictionary);
        string formContent = PdfEncoding.Latin1GetString(StreamDecoder.Decode(formStream.Dictionary, formStream.Data, objects));
        ImagePixelRewriteContentResult result = RewriteImagePixelsInContent(objects, formResources, formXObjects, formContent, targets, options, formTransform, referenceCounts, activeForms, removedMatches, ref nextObjectNumber);
        if (!string.Equals(result.Content, formContent, StringComparison.Ordinal)) {
            objects[formReference.ObjectNumber] = new PdfIndirectObject(formReference.ObjectNumber, formReference.Generation, new PdfStream(CleanStreamDictionary(formStream.Dictionary), PdfEncoding.Latin1GetBytes(result.Content)));
        }
        return result;
    }

    private static bool TryGetImageXObject(Dictionary<int, PdfIndirectObject> objects, PdfDictionary xObjects, string name, out PdfReference reference, out PdfStream stream) {
        if (xObjects.Items.TryGetValue(name, out PdfObject? value) &&
            value is PdfReference imageReference &&
            PdfObjectLookup.TryGet(objects, imageReference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfStream imageStream &&
            string.Equals(imageStream.Dictionary.Get<PdfName>("Subtype")?.Name, "Image", StringComparison.Ordinal)) {
            reference = imageReference;
            stream = imageStream;
            return true;
        }

        reference = default!;
        stream = default!;
        return false;
    }

    private static bool IsCurrentlySharedReference(
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, int> originalReferenceCounts,
        PdfReference reference) {
        return IsSharedReference(originalReferenceCounts, reference) ||
            IsSharedReference(CountIndirectReferenceUsage(objects), reference);
    }

    private static bool TryFindImageTarget(string resourceName, Matrix2D transform, ImageRedactionTarget[] targets, out ImageRedactionTarget target) {
        GetUnitRectangleBounds(transform, out double x, out double y, out double width, out double height);
        for (int i = 0; i < targets.Length; i++) {
            if (string.Equals(targets[i].ResourceName, resourceName, StringComparison.Ordinal) &&
                AreCloseImageCoordinate(targets[i].X, x) &&
                AreCloseImageCoordinate(targets[i].Y, y) &&
                AreCloseImageCoordinate(targets[i].Width, width) &&
                AreCloseImageCoordinate(targets[i].Height, height)) {
                target = targets[i];
                return true;
            }
        }

        target = default;
        return false;
    }

    private static ImageResourceInvocation[] ExtractImageResourceInvocations(string content) {
        var invocations = new List<ImageResourceInvocation>();
        Matrix2D ctm = Matrix2D.Identity;
        var stack = new Stack<Matrix2D>();
        var args = new List<ImageContentOperand>(8);
        int index = 0;
        int length = content.Length;

        while (index < length) {
            SkipWhiteSpace(content, ref index);
            if (index >= length) {
                break;
            }

            char current = content[index];
            if (current == '%') {
                SkipComment(content, ref index);
                continue;
            }

            if (current == '/') {
                args.Add(ReadNameOperand(content, ref index));
                continue;
            }

            if (current == '(') {
                SkipLiteralString(content, ref index);
                continue;
            }

            if (current == '<') {
                if (index + 1 < length && content[index + 1] == '<') {
                    SkipDictionary(content, ref index);
                } else {
                    SkipHexString(content, ref index);
                }

                continue;
            }

            if (current == '[') {
                SkipArray(content, ref index);
                continue;
            }

            if (current == ']' || current == '>') {
                index++;
                continue;
            }

            if (IsNumberStart(current)) {
                args.Add(ReadNumberOperand(content, ref index));
                continue;
            }

            string op = ReadOperator(content, ref index);
            if (op.Length == 0) {
                index++;
                continue;
            }

            switch (op) {
                case "q":
                    stack.Push(ctm);
                    args.Clear();
                    break;
                case "Q":
                    ctm = stack.Count > 0 ? stack.Pop() : Matrix2D.Identity;
                    args.Clear();
                    break;
                case "cm":
                    if (args.Count >= 6) {
                        ctm = Matrix2D.Multiply(ctm, new Matrix2D(
                            args[args.Count - 6].Number,
                            args[args.Count - 5].Number,
                            args[args.Count - 4].Number,
                            args[args.Count - 3].Number,
                            args[args.Count - 2].Number,
                            args[args.Count - 1].Number));
                    }

                    args.Clear();
                    break;
                case "Do":
                    if (args.Count >= 1 && !string.IsNullOrEmpty(args[args.Count - 1].Name)) {
                        ImageContentOperand operand = args[args.Count - 1];
                        invocations.Add(new ImageResourceInvocation(operand.Name!, ctm, operand.Start, operand.End));
                    }

                    args.Clear();
                    break;
                default:
                    args.Clear();
                    break;
            }
        }

        return invocations.ToArray();
    }

    private static string ReplaceInvocationResourceName(string content, ImageResourceInvocation invocation, string newName) {
        return content.Remove(invocation.NameStart, invocation.NameEnd - invocation.NameStart)
            .Insert(invocation.NameStart, "/" + PdfSyntaxEscaper.Name(newName));
    }

    private static string CreateUniqueResourceName(PdfDictionary xObjects, string baseName) {
        string normalized = string.IsNullOrWhiteSpace(baseName) ? "ImRedacted" : baseName + "Redacted";
        int index = 1;
        string candidate;
        do {
            candidate = normalized + index.ToString(CultureInfo.InvariantCulture);
            index++;
        } while (xObjects.Items.ContainsKey(candidate));

        return candidate;
    }

    private static bool CanRewriteImagePlacementPixels(Matrix2D transform) {
        return Math.Abs((transform.A * transform.D) - (transform.B * transform.C)) > ImagePixelRewriteTransformTolerance;
    }

    private static bool TryRewriteImageStreamPixels(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        string resourceName,
        PdfReference imageReference,
        PdfStream imageStream,
        ImageRedactionTarget target,
        Matrix2D transform,
        PdfRedactionApplyOptions options,
        ref int nextObjectNumber) {
        if (!TryGetSimpleWritableImage(imageStream, objects, options.MaximumDecodedImageBytes, out int width, out int height, out int components, out ImageSampleRewriteEncoder imageEncoder, out ImageSoftMaskRewriteTarget softMask)) {
            return TryRewriteNormalizedImagePixels(objects, resources, resourceName, imageReference, imageStream, target, transform, options, ref nextObjectNumber);
        }

        long expectedLengthLong = (long)width * height * components;
        if (expectedLengthLong <= 0 || expectedLengthLong > options.MaximumDecodedImageBytes || expectedLengthLong > int.MaxValue) {
            return false;
        }

        byte[] pixels = StreamDecoder.Decode(imageStream.Dictionary, imageStream.Data, objects, options.MaximumDecodedImageBytes);
        if (pixels.Length < expectedLengthLong) return false;

        if (!TryGetRedactionPixelBounds(target.Match.Area, transform, width, height, out int x0, out int y0, out int x1, out int y1)) {
            return false;
        }

        byte[] rewritten = (byte[])pixels.Clone();
        byte red = imageEncoder.EncodeSample(0, options.FillColor.R);
        byte green = imageEncoder.EncodeSample(1, options.FillColor.G);
        byte blue = imageEncoder.EncodeSample(2, options.FillColor.B);
        byte gray = imageEncoder.EncodeSample(0, ToGray(options.FillColor.R, options.FillColor.G, options.FillColor.B));

        for (int row = y0; row < y1; row++) {
            int rowOffset = row * width * components;
            for (int column = x0; column < x1; column++) {
                int offset = rowOffset + column * components;
                if (components == 1) {
                    rewritten[offset] = gray;
                } else {
                    rewritten[offset] = red;
                    rewritten[offset + 1] = green;
                    rewritten[offset + 2] = blue;
                }
            }
        }

        PdfDictionary dictionary = CleanStreamDictionary(imageStream.Dictionary);
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        if (softMask.HasMask &&
            TryRewriteSoftMaskPixels(objects, softMask, x0, y0, x1, y1, options.MaximumDecodedImageBytes, ref nextObjectNumber, out PdfReference rewrittenSoftMaskReference)) {
            dictionary.Items["SMask"] = rewrittenSoftMaskReference;
        } else if (softMask.HasMask) {
            return false;
        }

        byte[] compressed = CompressFlate(rewritten);
        objects[imageReference.ObjectNumber] = new PdfIndirectObject(imageReference.ObjectNumber, imageReference.Generation, new PdfStream(dictionary, compressed));
        return true;
    }

    private static bool TryRewriteNormalizedImagePixels(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary resources,
        string resourceName,
        PdfReference imageReference,
        PdfStream imageStream,
        ImageRedactionTarget target,
        Matrix2D transform,
        PdfRedactionApplyOptions options,
        ref int nextObjectNumber) {
        PdfExtractedImage extracted = ResourceResolver.BuildExtractedImage(
            0,
            resourceName,
            imageReference.ObjectNumber,
            0,
            imageStream,
            objects,
            resources: resources);
        if (!TryDecodeRedactionRaster(extracted, imageStream, objects, options, out int width, out int height, out byte[] rgba) ||
            !TryGetRedactionPixelBounds(target.Match.Area, transform, width, height, out int x0, out int y0, out int x1, out int y1)) {
            return false;
        }

        byte red = ToColorByte(options.FillColor.R);
        byte green = ToColorByte(options.FillColor.G);
        byte blue = ToColorByte(options.FillColor.B);
        for (int row = y0; row < y1; row++) {
            int rowOffset = row * width * 4;
            for (int column = x0; column < x1; column++) {
                int offset = rowOffset + column * 4;
                rgba[offset] = red;
                rgba[offset + 1] = green;
                rgba[offset + 2] = blue;
                rgba[offset + 3] = 255;
            }
        }

        byte[] rgb = new byte[checked(width * height * 3)];
        byte[] alpha = new byte[checked(width * height)];
        bool hasTransparency = false;
        for (int pixel = 0; pixel < alpha.Length; pixel++) {
            int source = pixel * 4;
            int destination = pixel * 3;
            rgb[destination] = rgba[source];
            rgb[destination + 1] = rgba[source + 1];
            rgb[destination + 2] = rgba[source + 2];
            alpha[pixel] = rgba[source + 3];
            hasTransparency = hasTransparency || alpha[pixel] != 255;
        }

        PdfDictionary dictionary = CleanStreamDictionary(imageStream.Dictionary);
        dictionary.Items.Remove("Decode");
        dictionary.Items.Remove("ImageMask");
        dictionary.Items.Remove("Mask");
        dictionary.Items.Remove("SMask");
        dictionary.Items["Width"] = new PdfNumber(width);
        dictionary.Items["Height"] = new PdfNumber(height);
        dictionary.Items["ColorSpace"] = new PdfName("DeviceRGB");
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        if (hasTransparency) {
            int maskObjectNumber = AllocateObjectNumber(objects, ref nextObjectNumber);
            dictionary.Items["SMask"] = new PdfReference(maskObjectNumber, 0);
            var maskDictionary = new PdfDictionary();
            maskDictionary.Items["Type"] = new PdfName("XObject");
            maskDictionary.Items["Subtype"] = new PdfName("Image");
            maskDictionary.Items["Width"] = new PdfNumber(width);
            maskDictionary.Items["Height"] = new PdfNumber(height);
            maskDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
            maskDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            maskDictionary.Items["Filter"] = new PdfName("FlateDecode");
            objects[maskObjectNumber] = new PdfIndirectObject(maskObjectNumber, 0, new PdfStream(maskDictionary, CompressFlate(alpha)));
        }

        objects[imageReference.ObjectNumber] = new PdfIndirectObject(imageReference.ObjectNumber, imageReference.Generation, new PdfStream(dictionary, CompressFlate(rgb)));
        return true;
    }

    private static bool TryDecodeRedactionRaster(
        PdfExtractedImage extracted,
        PdfStream imageStream,
        Dictionary<int, PdfIndirectObject> objects,
        PdfRedactionApplyOptions options,
        out int width,
        out int height,
        out byte[] rgba) {
        width = extracted.Width;
        height = extracted.Height;
        rgba = Array.Empty<byte>();
        long expectedLength = (long)width * height * 4;
        if (width <= 0 || height <= 0 || expectedLength > options.MaximumDecodedImageBytes || expectedLength > int.MaxValue) {
            return false;
        }

        if (extracted.IsImageFile &&
            string.Equals(extracted.FileExtension, "png", StringComparison.OrdinalIgnoreCase) &&
            OfficeRasterImageDecoder.TryDecode(extracted.Bytes, out OfficeRasterImage? raster) &&
            raster is not null) {
            width = raster.Width;
            height = raster.Height;
            rgba = raster.GetPixels();
        } else if (options.ImageDecoder is not null &&
            options.ImageDecoder.TryDecode(new PdfRedactionImageDecodeRequest(extracted), out PdfRedactionDecodedImage? decoded) &&
            decoded is not null) {
            width = decoded.Width;
            height = decoded.Height;
            rgba = decoded.GetRgbaPixels();
        } else {
            return false;
        }

        if (width != extracted.Width ||
            height != extracted.Height ||
            (long)width * height * 4 != rgba.Length ||
            rgba.Length > options.MaximumDecodedImageBytes) {
            rgba = Array.Empty<byte>();
            return false;
        }

        if (!extracted.HasUnresolvedTransparencyMask) {
            return true;
        }

        if (!string.Equals(extracted.TransparencyMaskKind, "explicit-mask-image", StringComparison.Ordinal) ||
            !imageStream.Dictionary.Items.TryGetValue("Mask", out PdfObject? maskObject) ||
            PdfObjectLookup.Resolve(objects, maskObject) is not PdfStream maskStream ||
            !PdfImageMaskNormalizer.TryBuildPngFile(width, height, maskStream, objects, out byte[] maskPng) ||
            !OfficeRasterImageDecoder.TryDecode(maskPng, out OfficeRasterImage? maskRaster) ||
            maskRaster is null ||
            maskRaster.Width != width ||
            maskRaster.Height != height) {
            rgba = Array.Empty<byte>();
            return false;
        }

        byte[] maskPixels = maskRaster.GetPixels();
        for (int pixel = 0; pixel < width * height; pixel++) {
            rgba[pixel * 4 + 3] = maskPixels[pixel * 4 + 3];
        }

        return true;
    }

    private static int AllocateObjectNumber(Dictionary<int, PdfIndirectObject> objects, ref int nextObjectNumber) {
        int objectNumber = nextObjectNumber++;
        while (objects.ContainsKey(objectNumber)) objectNumber = nextObjectNumber++;
        return objectNumber;
    }

    private static bool TryGetSimpleWritableImage(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumDecodedImageBytes,
        out int width,
        out int height,
        out int components,
        out ImageSampleRewriteEncoder imageEncoder,
        out ImageSoftMaskRewriteTarget softMask) {
        width = (int)(stream.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        height = (int)(stream.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        components = 0;
        imageEncoder = default;
        softMask = default;

        if (width <= 0 ||
            height <= 0 ||
            (int)(stream.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0) != 8 ||
            HasTrueBoolean(stream.Dictionary, "ImageMask") ||
            stream.Dictionary.Items.ContainsKey("Mask") ||
            StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) {
            return false;
        }

        string colorSpace = ReadSimpleColorSpaceName(stream.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ? colorSpaceObject : null, objects);
        if (string.Equals(colorSpace, "DeviceGray", StringComparison.Ordinal)) {
            components = 1;
        } else if (string.Equals(colorSpace, "DeviceRGB", StringComparison.Ordinal)) {
            components = 3;
        } else {
            return false;
        }

        if (!TryCreateImageSampleRewriteEncoder(stream.Dictionary, components, objects, out imageEncoder) ||
            !TryGetWritableSoftMask(stream, objects, width, height, maximumDecodedImageBytes, out softMask)) {
            return false;
        }

        return true;
    }

    private static bool TryGetWritableSoftMask(
        PdfStream imageStream,
        Dictionary<int, PdfIndirectObject> objects,
        int width,
        int height,
        int maximumDecodedImageBytes,
        out ImageSoftMaskRewriteTarget softMask) {
        softMask = default;
        if (!imageStream.Dictionary.Items.TryGetValue("SMask", out PdfObject? softMaskObject)) {
            return true;
        }

        if (PdfObjectLookup.Resolve(objects, softMaskObject) is PdfName softMaskName &&
            string.Equals(softMaskName.Name, "None", StringComparison.Ordinal)) {
            return true;
        }

        if (softMaskObject is not PdfReference softMaskReference ||
            !PdfObjectLookup.TryGet(objects, softMaskReference, out PdfIndirectObject? softMaskIndirect) ||
            softMaskIndirect.Value is not PdfStream softMaskStream) {
            return false;
        }

        if ((int)(softMaskStream.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0) != width ||
            (int)(softMaskStream.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0) != height ||
            (int)(softMaskStream.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0) != 8 ||
            !string.Equals(ReadSimpleColorSpaceName(softMaskStream.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceObject) ? colorSpaceObject : null, objects), "DeviceGray", StringComparison.Ordinal) ||
            HasTrueBoolean(softMaskStream.Dictionary, "ImageMask") ||
            softMaskStream.Dictionary.Items.ContainsKey("Mask") ||
            softMaskStream.Dictionary.Items.ContainsKey("SMask") ||
            StreamDecoder.GetUnsupportedFilters(softMaskStream.Dictionary, objects).Count != 0) {
            return false;
        }

        if (!TryCreateImageSampleRewriteEncoder(softMaskStream.Dictionary, 1, objects, out ImageSampleRewriteEncoder maskEncoder)) {
            return false;
        }

        long expectedLengthLong = (long)width * height;
        if (expectedLengthLong <= 0 || expectedLengthLong > maximumDecodedImageBytes || expectedLengthLong > int.MaxValue) {
            return false;
        }

        byte[] maskPixels = StreamDecoder.Decode(softMaskStream.Dictionary, softMaskStream.Data, objects, maximumDecodedImageBytes);
        if (maskPixels.Length < expectedLengthLong) return false;

        softMask = new ImageSoftMaskRewriteTarget(softMaskReference, softMaskStream, width, height, maskEncoder);
        return true;
    }

    private static bool TryRewriteSoftMaskPixels(
        Dictionary<int, PdfIndirectObject> objects,
        ImageSoftMaskRewriteTarget softMask,
        int x0,
        int y0,
        int x1,
        int y1,
        int maximumDecodedImageBytes,
        ref int nextObjectNumber,
        out PdfReference rewrittenSoftMaskReference) {
        rewrittenSoftMaskReference = default!;
        if (!softMask.HasMask) {
            return false;
        }

        long expectedLengthLong = (long)softMask.Width * softMask.Height;
        if (expectedLengthLong <= 0 || expectedLengthLong > maximumDecodedImageBytes || expectedLengthLong > int.MaxValue) {
            return false;
        }

        byte[] pixels = StreamDecoder.Decode(softMask.Stream.Dictionary, softMask.Stream.Data, objects, maximumDecodedImageBytes);
        if (pixels.Length < expectedLengthLong) return false;

        byte[] rewritten = (byte[])pixels.Clone();
        for (int row = y0; row < y1; row++) {
            int rowOffset = row * softMask.Width;
            for (int column = x0; column < x1; column++) {
                rewritten[rowOffset + column] = softMask.Encoder.EncodeSample(0, 1D);
            }
        }

        int objectNumber = nextObjectNumber++;
        while (objects.ContainsKey(objectNumber)) {
            objectNumber = nextObjectNumber++;
        }

        rewrittenSoftMaskReference = new PdfReference(objectNumber, 0);
        PdfDictionary dictionary = CleanStreamDictionary(softMask.Stream.Dictionary);
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, new PdfStream(dictionary, CompressFlate(rewritten)));
        return true;
    }

    private static bool TryGetRedactionPixelBounds(PdfRedactionArea area, Matrix2D transform, int width, int height, out int x0, out int y0, out int x1, out int y1) {
        double determinant = (transform.A * transform.D) - (transform.B * transform.C);
        if (Math.Abs(determinant) <= ImagePixelRewriteTransformTolerance) {
            x0 = y0 = x1 = y1 = 0;
            return false;
        }

        double leftFraction = double.PositiveInfinity;
        double rightFraction = double.NegativeInfinity;
        double bottomFraction = double.PositiveInfinity;
        double topFraction = double.NegativeInfinity;
        AddInversePoint(area.X, area.Y);
        AddInversePoint(area.X + area.Width, area.Y);
        AddInversePoint(area.X, area.Y + area.Height);
        AddInversePoint(area.X + area.Width, area.Y + area.Height);
        leftFraction = ClampUnit(leftFraction);
        rightFraction = ClampUnit(rightFraction);
        bottomFraction = ClampUnit(bottomFraction);
        topFraction = ClampUnit(topFraction);
        if (rightFraction <= leftFraction || topFraction <= bottomFraction) {
            x0 = y0 = x1 = y1 = 0;
            return false;
        }

        x0 = ClampPixel((int)Math.Floor(leftFraction * width), width);
        x1 = ClampPixel((int)Math.Ceiling(rightFraction * width), width);
        y0 = ClampPixel((int)Math.Floor((1D - topFraction) * height), height);
        y1 = ClampPixel((int)Math.Ceiling((1D - bottomFraction) * height), height);
        return x1 > x0 && y1 > y0;

        void AddInversePoint(double pageX, double pageY) {
            double translatedX = pageX - transform.E;
            double translatedY = pageY - transform.F;
            double unitX = ((transform.D * translatedX) - (transform.C * translatedY)) / determinant;
            double unitY = ((-transform.B * translatedX) + (transform.A * translatedY)) / determinant;
            leftFraction = Math.Min(leftFraction, unitX);
            rightFraction = Math.Max(rightFraction, unitX);
            bottomFraction = Math.Min(bottomFraction, unitY);
            topFraction = Math.Max(topFraction, unitY);
        }
    }

    private static PdfDictionary EnsureFormResources(Dictionary<int, PdfIndirectObject> objects, PdfStream formStream) {
        if (formStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)) {
            PdfDictionary resolved = ResolveDictionary(objects, resourcesObject) ?? new PdfDictionary();
            if (resourcesObject is PdfReference) {
                resolved = CloneDictionary(resolved);
                formStream.Dictionary.Items["Resources"] = resolved;
            }

            return resolved;
        }

        var resources = new PdfDictionary();
        formStream.Dictionary.Items["Resources"] = resources;
        return resources;
    }

    private static PdfDictionary EnsureResourceXObjects(Dictionary<int, PdfIndirectObject> objects, PdfDictionary resources) {
        if (resources.Items.TryGetValue("XObject", out PdfObject? xObjectObject)) {
            PdfDictionary resolved = ResolveDictionary(objects, xObjectObject) ?? new PdfDictionary();
            if (xObjectObject is PdfReference) {
                resolved = CloneDictionary(resolved);
                resources.Items["XObject"] = resolved;
            }

            return resolved;
        }

        var xObjects = new PdfDictionary();
        resources.Items["XObject"] = xObjects;
        return xObjects;
    }

    private static string ReadSimpleColorSpaceName(PdfObject? colorSpaceObject, Dictionary<int, PdfIndirectObject> objects) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, colorSpaceObject);
        if (resolved is PdfName name) {
            return name.Name;
        }

        if (resolved is PdfArray array &&
            array.Items.Count > 0 &&
            PdfObjectLookup.Resolve(objects, array.Items[0]) is PdfName arrayName) {
            return arrayName.Name;
        }

        return string.Empty;
    }

    private static bool HasTrueBoolean(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            value is PdfBoolean boolean &&
            boolean.Value;
    }

    private static double ClampUnit(double value) {
        if (value <= 0D) {
            return 0D;
        }

        return value >= 1D ? 1D : value;
    }

    private static int ClampPixel(int value, int length) {
        if (value <= 0) {
            return 0;
        }

        return value >= length ? length : value;
    }

    private static bool TryCreateImageSampleRewriteEncoder(
        PdfDictionary dictionary,
        int componentCount,
        Dictionary<int, PdfIndirectObject> objects,
        out ImageSampleRewriteEncoder encoder) {
        encoder = default;
        if (componentCount <= 0) {
            return false;
        }

        if (!dictionary.Items.TryGetValue("Decode", out PdfObject? decodeObject)) {
            encoder = ImageSampleRewriteEncoder.CreateIdentity(componentCount);
            return true;
        }

        if (PdfObjectLookup.Resolve(objects, decodeObject) is not PdfArray decodeArray ||
            decodeArray.Items.Count < componentCount * 2) {
            return false;
        }

        double[] minimums = new double[componentCount];
        double[] maximums = new double[componentCount];
        for (int component = 0; component < componentCount; component++) {
            if (PdfObjectLookup.Resolve(objects, decodeArray.Items[component * 2]) is not PdfNumber minimum ||
                PdfObjectLookup.Resolve(objects, decodeArray.Items[component * 2 + 1]) is not PdfNumber maximum ||
                double.IsNaN(minimum.Value) ||
                double.IsInfinity(minimum.Value) ||
                double.IsNaN(maximum.Value) ||
                double.IsInfinity(maximum.Value) ||
                Math.Abs(maximum.Value - minimum.Value) <= double.Epsilon) {
                return false;
            }

            minimums[component] = minimum.Value;
            maximums[component] = maximum.Value;
        }

        encoder = new ImageSampleRewriteEncoder(minimums, maximums);
        return true;
    }

    private static byte ToColorByte(double value) {
        if (value <= 0D) {
            return 0;
        }

        if (value >= 1D) {
            return 255;
        }

        return (byte)Math.Round(value * 255D);
    }

    private static double ToGray(double red, double green, double blue) {
        return (red * 0.299D) + (green * 0.587D) + (blue * 0.114D);
    }

    private static byte[] CompressFlate(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        uint adler = Adler32(data);
        output.WriteByte((byte)((adler >> 24) & 0xFF));
        output.WriteByte((byte)((adler >> 16) & 0xFF));
        output.WriteByte((byte)((adler >> 8) & 0xFF));
        output.WriteByte((byte)(adler & 0xFF));
        return output.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private readonly struct ImagePixelRewriteContentResult {
        public ImagePixelRewriteContentResult(bool hasChanges, string content) {
            HasChanges = hasChanges;
            Content = content;
        }

        public bool HasChanges { get; }

        public string Content { get; }
    }

    private readonly struct ImageResourceInvocation {
        public ImageResourceInvocation(string name, Matrix2D transform, int nameStart, int nameEnd) {
            Name = name;
            Transform = transform;
            NameStart = nameStart;
            NameEnd = nameEnd;
        }

        public string Name { get; }

        public Matrix2D Transform { get; }

        public int NameStart { get; }

        public int NameEnd { get; }
    }

    private readonly struct ImageSampleRewriteEncoder {
        private readonly double[]? _minimums;
        private readonly double[]? _maximums;

        public ImageSampleRewriteEncoder(double[] minimums, double[] maximums) {
            _minimums = minimums;
            _maximums = maximums;
        }

        public static ImageSampleRewriteEncoder CreateIdentity(int componentCount) {
            double[] minimums = new double[componentCount];
            double[] maximums = new double[componentCount];
            for (int i = 0; i < maximums.Length; i++) {
                maximums[i] = 1D;
            }

            return new ImageSampleRewriteEncoder(minimums, maximums);
        }

        public byte EncodeSample(int componentIndex, double decodedValue) {
            if (_minimums is null || _maximums is null || _minimums.Length == 0 || _maximums.Length == 0) {
                return ToColorByte(decodedValue);
            }

            int safeComponentIndex = Math.Min(componentIndex, _minimums.Length - 1);
            double minimum = _minimums[safeComponentIndex];
            double maximum = _maximums[safeComponentIndex];
            double encodedValue = (decodedValue - minimum) / (maximum - minimum);
            return ToColorByte(encodedValue);
        }
    }

    private readonly struct ImageSoftMaskRewriteTarget {
        public ImageSoftMaskRewriteTarget(PdfReference reference, PdfStream stream, int width, int height, ImageSampleRewriteEncoder encoder) {
            Reference = reference;
            Stream = stream;
            Width = width;
            Height = height;
            Encoder = encoder;
            HasMask = true;
        }

        public bool HasMask { get; }

        public PdfReference Reference { get; }

        public PdfStream Stream { get; }

        public int Width { get; }

        public int Height { get; }

        public ImageSampleRewriteEncoder Encoder { get; }
    }
}
