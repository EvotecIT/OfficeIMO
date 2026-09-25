using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    /// <summary>Adds a visual Stamp annotation to an existing page using the shared full-rewrite or append-only planner.</summary>
    public static PdfAnnotationEditResult AddStampAnnotation(byte[] pdf, PdfStampAnnotationOptions? options = null) =>
        AddStampAnnotation(pdf, options, readOptions: null);

    /// <summary>Adds a visual Stamp annotation using explicit read limits or credentials.</summary>
    public static PdfAnnotationEditResult AddStampAnnotation(byte[] pdf, PdfStampAnnotationOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfStampAnnotationOptions effective = options ?? new PdfStampAnnotationOptions();
        ValidateStampOptions(effective);
        byte[]? imageBytes = effective.GetImageBytes();
        PdfWriter.PdfImageStream? imageStream = null;
        if (imageBytes is not null) {
            PdfImageInput.EnsureWithinLimit(imageBytes.LongLength, effective.MaximumEncodedImageBytes);
            PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(imageBytes);
            if (!PdfWriter.TryBuildImageStream(prepared, effective.Width, effective.Height, out PdfWriter.PdfImageStream builtImage, out string? reason)) {
                throw new NotSupportedException(reason ?? "Stamp image format is not supported.");
            }
            imageStream = builtImage;
        }
        PdfMutationPlan mutationPlan = PdfMutationPlanner.Require(
            pdf,
            PdfMutationOperation.ModifyAnnotations,
            readOptions,
            executionPreference: effective.ExecutionPreference);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0) {
            throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        }

        List<int> pages = GetPageObjectNumbersInDocumentOrder(objects);
        if (effective.PageNumber > pages.Count) {
            throw new ArgumentOutOfRangeException(nameof(options), "Stamp page number exceeds the PDF page count.");
        }

        int pageObjectNumber = pages[effective.PageNumber - 1];
        PdfIndirectObject pageObject = objects[pageObjectNumber];
        if (pageObject.Value is not PdfDictionary page) {
            throw new InvalidOperationException("The selected page object is not a dictionary.");
        }

        int nextObjectNumber = NextAnnotationObjectNumber(objects);
        // Reserve every generated object before mutating the graph; post-increment must not wrap.
        _ = checked(nextObjectNumber + (imageStream?.SoftMask is not null ? 4 : 3));
        var generatedObjects = new List<int>();
        int? fontObjectNumber = null;
        int? imageObjectNumber = null;
        if (imageStream is null) {
            fontObjectNumber = nextObjectNumber++;
            objects[fontObjectNumber.Value] = new PdfIndirectObject(fontObjectNumber.Value, 0, BuildStampFont());
            generatedObjects.Add(fontObjectNumber.Value);
        } else {
            int? softMaskObjectNumber = null;
            if (imageStream.SoftMask is not null) {
                softMaskObjectNumber = nextObjectNumber++;
                objects[softMaskObjectNumber.Value] = new PdfIndirectObject(softMaskObjectNumber.Value, 0, PdfWriter.BuildImageXObject(imageStream.SoftMask));
                generatedObjects.Add(softMaskObjectNumber.Value);
            }
            imageObjectNumber = nextObjectNumber++;
            objects[imageObjectNumber.Value] = new PdfIndirectObject(imageObjectNumber.Value, 0, PdfWriter.BuildImageXObject(imageStream, softMaskObjectNumber));
            generatedObjects.Add(imageObjectNumber.Value);
        }
        int appearanceObjectNumber = nextObjectNumber++;
        int annotationObjectNumber = nextObjectNumber++;
        PdfStream appearance = BuildStampAppearance(effective, fontObjectNumber, imageObjectNumber, imageStream);
        objects[appearanceObjectNumber] = new PdfIndirectObject(
            appearanceObjectNumber,
            0,
            appearance);
        objects[annotationObjectNumber] = new PdfIndirectObject(
            annotationObjectNumber,
            0,
            BuildStampAnnotation(effective, pageObject, appearanceObjectNumber));
        generatedObjects.Add(appearanceObjectNumber);
        generatedObjects.Add(annotationObjectNumber);

        int annotationsOwnerObjectNumber = AddAnnotationReference(
            objects,
            pageObjectNumber,
            page,
            new PdfReference(annotationObjectNumber, 0));
        PdfGeneratedOutputGrowth generatedGrowth = BuildGeneratedOutputGrowth(
            objects,
            generatedObjects.Concat(new[] { annotationsOwnerObjectNumber }),
            additionalAnnotationsPerPage: 1,
            additionalRevisions: mutationPlan.ExecutionMode == PdfMutationExecutionMode.AppendOnly ? 1 : 0);
        if (mutationPlan.ExecutionMode == PdfMutationExecutionMode.AppendOnly) {
            int[] changedObjectNumbers = generatedObjects.Concat(new[] { annotationsOwnerObjectNumber }).Distinct().ToArray();
            byte[] appended = PdfIncrementalObjectWriter.Append(
                pdf,
                objects,
                mutationPlan.Preflight.Probe.Security,
                trailerRaw,
                changedObjectNumbers,
                encryptionHandler: GetAppendEncryptionHandler(objects, trailerRaw, readOptions, mutationPlan.Preflight.Probe.Security));
            PdfLoadOptions outputReadOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, pdf, appended, generatedGrowth);
            PdfSignatureMutationReport proof = BuildAppendOnlyProof(pdf, appended, mutationPlan, readOptions, outputReadOptions);
            return new PdfAnnotationEditResult(appended, 1, mutationPlan, proof, readOptions: outputReadOptions);
        }

        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber);
        byte[] rewritten = RewriteAllObjects(objects, catalogObjectNumber, PdfReadDocument.Open(pdf, readOptions).UncheckedMetadata, pdf, out var numberMap);
        return CreateFullRewriteResult(pdf, rewritten, 1, mutationPlan, annotationsChanged: true, readOptions: readOptions, generatedGrowth: generatedGrowth, objectNumberMap: numberMap);
    }

    private static int AddAnnotationReference(
        Dictionary<int, PdfIndirectObject> objects,
        int pageObjectNumber,
        PdfDictionary page,
        PdfReference annotationReference) {
        if (!page.Items.TryGetValue("Annots", out PdfObject? annotsObject)) {
            var annotations = new PdfArray();
            annotations.Items.Add(annotationReference);
            page.Items["Annots"] = annotations;
            return pageObjectNumber;
        }

        if (PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray existingAnnotations) {
            throw new NotSupportedException("The selected page /Annots entry is not a readable array.");
        }

        existingAnnotations.Items.Add(annotationReference);
        return annotsObject is PdfReference reference ? reference.ObjectNumber : pageObjectNumber;
    }

    private static PdfDictionary BuildStampFont() {
        var font = new PdfDictionary();
        font.Items["Type"] = new PdfName("Font");
        font.Items["Subtype"] = new PdfName("Type1");
        font.Items["BaseFont"] = new PdfName("Helvetica");
        font.Items["Encoding"] = new PdfName("WinAnsiEncoding");
        return font;
    }

    private static PdfStream BuildStampAppearance(PdfStampAnnotationOptions options, int? fontObjectNumber, int? imageObjectNumber, PdfWriter.PdfImageStream? imageStream) {
        var resources = new PdfDictionary();
        if (fontObjectNumber.HasValue) {
            var fontResources = new PdfDictionary();
            fontResources.Items["Helv"] = new PdfReference(fontObjectNumber.Value, 0);
            resources.Items["Font"] = fontResources;
        }
        if (imageObjectNumber.HasValue) {
            var xObjects = new PdfDictionary();
            xObjects.Items["Im1"] = new PdfReference(imageObjectNumber.Value, 0);
            resources.Items["XObject"] = xObjects;
        }
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["FormType"] = new PdfNumber(1);
        dictionary.Items["BBox"] = BuildNumberArray(0D, 0D, options.Width, options.Height);
        dictionary.Items["Resources"] = resources;
        string content;
        if (imageStream is not null) {
            OfficeImagePlacement placement = OfficeImageRenderPlan.CreateBottomLeft(
                imageStream.PixelWidth, imageStream.PixelHeight, 0D, 0D, options.Width, options.Height, OfficeImageFit.Contain).ImagePlacement;
            content = "q\n" + FormatStampNumber(placement.Width) + " 0 0 " +
                FormatStampNumber(placement.Height) + " " + FormatStampNumber(placement.X) + " " +
                FormatStampNumber(placement.Y) + " cm\n/Im1 Do\nQ\n";
        } else {
            content = PdfAnnotationDictionaryBuilder.BuildStampAppearanceContent(
                options.Width,
                options.Height,
                options.StampName,
                options.StrokeColor,
                options.FillColor,
                options.BorderWidth);
        }
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static string FormatStampNumber(double value) => value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

    private static PdfDictionary BuildStampAnnotation(
        PdfStampAnnotationOptions options,
        PdfIndirectObject pageObject,
        int appearanceObjectNumber) {
        var normalAppearance = new PdfDictionary();
        normalAppearance.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
        var annotation = new PdfDictionary();
        annotation.Items["Type"] = new PdfName("Annot");
        annotation.Items["Subtype"] = new PdfName("Stamp");
        annotation.Items["Rect"] = BuildNumberArray(
            options.X,
            options.Y,
            options.X + options.Width,
            options.Y + options.Height);
        annotation.Items["Name"] = new PdfName(options.StampName);
        annotation.Items["F"] = new PdfNumber(options.Flags);
        annotation.Items["C"] = BuildNumberArray(options.StrokeColor.R, options.StrokeColor.G, options.StrokeColor.B);
        annotation.Items["P"] = new PdfReference(pageObject.ObjectNumber, pageObject.Generation);
        annotation.Items["AP"] = normalAppearance;
        AddOptionalText(annotation, "Contents", options.Contents);
        AddOptionalText(annotation, "T", options.Title);
        AddOptionalText(annotation, "NM", options.Name);
        return annotation;
    }

    private static void AddOptionalText(PdfDictionary dictionary, string key, string? value) {
        if (value is not null) {
            dictionary.Items[key] = new PdfStringObj(value, useTextStringEncoding: true);
        }
    }

    private static PdfArray BuildNumberArray(params double[] values) {
        var array = new PdfArray();
        for (int i = 0; i < values.Length; i++) {
            array.Items.Add(new PdfNumber(values[i]));
        }

        return array;
    }

    private static int NextAnnotationObjectNumber(Dictionary<int, PdfIndirectObject> objects) {
        return objects.Count == 0 ? 1 : checked(objects.Keys.Max() + 1);
    }

    private static void ValidateStampOptions(PdfStampAnnotationOptions options) {
        if (options.PageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "Stamp page number must be positive.");
        }

        ValidateFinite(options.X, nameof(options.X));
        ValidateFinite(options.Y, nameof(options.Y));
        Guard.Positive(options.Width, nameof(options.Width));
        Guard.Positive(options.Height, nameof(options.Height));
        Guard.NotNullOrWhiteSpace(options.StampName, nameof(options.StampName));
        Guard.NonNegative(options.Flags, nameof(options.Flags));
        Guard.NonNegative(options.BorderWidth, nameof(options.BorderWidth));
        PdfImageInput.ValidateMaximum(options.MaximumEncodedImageBytes);
    }

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Stamp coordinates must be finite.");
        }
    }
}
