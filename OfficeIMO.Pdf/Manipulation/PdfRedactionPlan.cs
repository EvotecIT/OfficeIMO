namespace OfficeIMO.Pdf;

/// <summary>Preview of text, image placements, and annotations that intersect requested redaction rectangles.</summary>
public sealed class PdfRedactionPlan {
    internal PdfRedactionPlan(
        PdfDocumentPreflight preflight,
        IReadOnlyList<PdfRedactionArea> areas,
        IReadOnlyList<PdfRedactionMatch> matches,
        IReadOnlyList<PdfDiagnosticFinding> findings,
        IReadOnlyList<string>? searchCriteria,
        string sourceSha256,
        IReadOnlyList<string>? pageIdentities = null,
        IReadOnlyList<IReadOnlyList<PdfRedactionTextObjectScope>>? reviewedTextObjectScopes = null) {
        Preflight = preflight;
        Areas = areas;
        Matches = matches;
        Findings = findings;
        SearchCriteria = searchCriteria ?? Array.Empty<string>();
        SourceSha256 = sourceSha256;
        PageIdentities = pageIdentities ?? Array.Empty<string>();
        ReviewedTextObjectScopes = reviewedTextObjectScopes ?? Array.Empty<IReadOnlyList<PdfRedactionTextObjectScope>>();
    }

    /// <summary>Preflight result used while creating the plan.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Requested redaction areas.</summary>
    public IReadOnlyList<PdfRedactionArea> Areas { get; }

    /// <summary>Text blocks, image placements, and annotations intersecting the requested areas.</summary>
    public IReadOnlyList<PdfRedactionMatch> Matches { get; }

    /// <summary>Diagnostics and warnings for the plan.</summary>
    public IReadOnlyList<PdfDiagnosticFinding> Findings { get; }

    /// <summary>Stable descriptions of literal, regex, logical-kind, or form-field criteria used to derive the areas.</summary>
    public IReadOnlyList<string> SearchCriteria { get; }

    /// <summary>SHA-256 fingerprint of the exact PDF bytes inspected while creating this plan.</summary>
    public string SourceSha256 { get; }

    internal IReadOnlyList<string> PageIdentities { get; }

    internal IReadOnlyList<IReadOnlyList<PdfRedactionTextObjectScope>> ReviewedTextObjectScopes { get; }

    /// <summary>True when the source was inspectable and the plan contains no blocking findings.</summary>
    public bool IsReviewable =>
        Preflight.CanReadLogicalObjects &&
        Findings.All(static finding => finding.Severity != PdfDiagnosticSeverity.Error);

    /// <summary>True when the plan areas were derived from explicit search criteria.</summary>
    public bool IsSearchDriven => SearchCriteria.Count > 0;

    /// <summary>True when at least one match was found.</summary>
    public bool HasMatches => Matches.Count > 0;

    internal bool MatchesSource(byte[] pdf) =>
        string.Equals(SourceSha256, ComputeSourceSha256(pdf), StringComparison.Ordinal);

    internal static string ComputeSourceSha256(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
#if NET6_0_OR_GREATER
        return Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(pdf));
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(pdf));
#endif
    }

    internal static IReadOnlyList<string> CapturePageIdentities(
        PdfReadDocument document,
        IReadOnlyList<PdfRedactionArea> reviewedAreas) {
        IReadOnlyList<IReadOnlyList<PdfRedactionTextObjectScope>> reviewedTextObjectScopes = CaptureReviewedTextObjectScopes(document, reviewedAreas);
        return CapturePageIdentities(document, reviewedAreas, reviewedTextObjectScopes);
    }

    internal static IReadOnlyList<string> CapturePageIdentities(
        PdfReadDocument document,
        IReadOnlyList<PdfRedactionArea> reviewedAreas,
        IReadOnlyList<IReadOnlyList<PdfRedactionTextObjectScope>> reviewedTextObjectScopes) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(reviewedAreas, nameof(reviewedAreas));
        Guard.NotNull(reviewedTextObjectScopes, nameof(reviewedTextObjectScopes));
        IReadOnlyDictionary<int, string> stablePageReferences = CreateStablePageReferenceLabels(document);
        var identities = new string[document.Pages.Count];
        for (int i = 0; i < document.Pages.Count; i++) {
            PdfReadPage page = document.Pages[i];
            PdfPageGeometry geometry = page.GetGeometry();
            int pageNumber = i + 1;
            PdfRedactionArea[] pageAreas = reviewedAreas
                .Where(area => area.PageNumber == pageNumber)
                .ToArray();
            var identity = new System.Text.StringBuilder();
            IReadOnlyList<PdfPageDrawingEffectTransition> drawingEffects = page.GetIdentityGraphicsEffectTransitions();
            identity.Append(string.Join("|", new[] {
                page.GetRotationDegrees().ToString(System.Globalization.CultureInfo.InvariantCulture),
                FormatPageBoxIdentity(geometry.MediaBox),
                FormatPageBoxIdentity(geometry.CropBox),
                geometry.UserUnit?.ToString("R", System.Globalization.CultureInfo.InvariantCulture) ?? "null"
            }));
            IReadOnlyList<PdfRedactionTextObjectScope> pageReviewedTextObjectScopes = i < reviewedTextObjectScopes.Count
                ? reviewedTextObjectScopes[i]
                : Array.Empty<PdfRedactionTextObjectScope>();
            AppendUnredactedTextIdentity(identity, document, page, pageAreas, pageReviewedTextObjectScopes, drawingEffects);
            AppendUnredactedPathIdentity(identity, document, page, pageAreas, drawingEffects);
            AppendUnredactedImageIdentity(identity, document, page, pageNumber, pageAreas, drawingEffects);
            AppendUnredactedAnnotationIdentity(identity, document, page, pageAreas, stablePageReferences);
            AppendUnredactedLinkIdentity(identity, page, pageAreas);
            AppendPageRenderingResourceIdentity(identity, document, page);
            identity.Append("|C:OCProperties:");
            PdfRedactionImageIdentity.AppendObjectGraph(
                identity,
                document.CatalogDictionary?.Items.TryGetValue("OCProperties", out PdfObject? optionalContent) == true ? optionalContent : null,
                document.Objects);
            identities[i] = ComputeIdentityHash(identity.ToString());
        }

        return identities;
    }

    internal static IReadOnlyList<IReadOnlyList<PdfRedactionTextObjectScope>> CaptureReviewedTextObjectScopes(
        PdfReadDocument document,
        IReadOnlyList<PdfRedactionArea> reviewedAreas) {
        var result = new IReadOnlyList<PdfRedactionTextObjectScope>[document.Pages.Count];
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            int pageNumber = pageIndex + 1;
            PdfRedactionArea[] pageAreas = reviewedAreas.Where(area => area.PageNumber == pageNumber).ToArray();
            IReadOnlyList<PdfTextSpan> spans = document.Pages[pageIndex].GetTextSpansIncludingHiddenOptionalContent();
            result[pageIndex] = CreateTextObjectScopes(spans, pageAreas)
                .Where(static scope => scope.HasReviewedIntersection)
                .ToArray();
        }
        return result;
    }

    private static Dictionary<int, string> CreateStablePageReferenceLabels(PdfReadDocument document) {
        var labels = new Dictionary<int, string>();
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            int pageNumber = pageIndex + 1;
            labels[document.Pages[pageIndex].ObjectNumber] = "page:" + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        return labels;
    }

    private static void AppendUnredactedTextIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas,
        IReadOnlyList<PdfRedactionTextObjectScope> reviewedTextObjectScopes,
        IReadOnlyList<PdfPageDrawingEffectTransition> drawingEffects) {
        IReadOnlyList<PdfTextSpan> spans = page.GetTextSpansIncludingHiddenOptionalContent();
        var ignoredTextObjectKeys = new HashSet<PdfContentOrderKey>();
        PdfRedactionTextObjectScope[] currentTextObjectScopes = CreateTextObjectScopes(spans);
        for (int currentIndex = 0; currentIndex < currentTextObjectScopes.Length; currentIndex++) {
            PdfRedactionTextObjectScope current = currentTextObjectScopes[currentIndex];
            if (reviewedTextObjectScopes.Any(reviewed => reviewed.Matches(current))) ignoredTextObjectKeys.Add(current.Key);
        }
        for (int i = 0; i < spans.Count; i++) {
            PdfTextSpan span = spans[i];
            PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(span);
            if (span.TextObjectOrderKey != null && ignoredTextObjectKeys.Contains(span.TextObjectOrderKey) ||
                span.TextObjectOrderKey == null && IntersectsReviewedArea(pageAreas, bounds.Left, bounds.Bottom, bounds.Width, bounds.Height)) continue;
            identity.Append("|T:")
                .Append(span.Text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(':').Append(span.Text)
                .Append(':').Append(FormatIdentityNumber(span.X))
                .Append(',').Append(FormatIdentityNumber(span.Y))
                .Append(',').Append(FormatIdentityNumber(span.Advance))
                .Append(',').Append(FormatIdentityNumber(span.FontSize))
                .Append(',').Append(FormatIdentityNumber(span.RotationDegrees))
                .Append(',').Append(span.IsVisible ? '1' : '0')
                .Append(',').Append(span.TextRenderingMode);
            if (span.TextToPageTransform.HasValue) {
                Matrix2D transform = span.TextToPageTransform.Value;
                identity.Append(":tm:").Append(FormatIdentityNumber(transform.A))
                    .Append(',').Append(FormatIdentityNumber(transform.B))
                    .Append(',').Append(FormatIdentityNumber(transform.C))
                    .Append(',').Append(FormatIdentityNumber(transform.D))
                    .Append(',').Append(FormatIdentityNumber(transform.E))
                    .Append(',').Append(FormatIdentityNumber(transform.F));
            }
            AppendIdentityString(identity, span.BaseFont ?? span.DrawingFontFamily ?? span.FontResource);
            AppendIdentityColor(identity, span.Color);
            AppendIdentityString(identity, span.VisualPaintIdentity);
            PdfRedactionImageIdentity.AppendClip(identity, span.ClipPath);
            AppendDrawingEffectIdentity(identity, document, PdfReadPage.ResolveDrawingEffect(drawingEffects, span.PaintOrder, contentOrderKey: span.ContentOrderKey));
        }
    }

    private static PdfRedactionTextObjectScope[] CreateTextObjectScopes(
        IReadOnlyList<PdfTextSpan> spans,
        IReadOnlyList<PdfRedactionArea>? reviewedAreas = null) =>
        spans
            .Where(static span => span.TextObjectOrderKey is not null)
            .GroupBy(static span => span.TextObjectOrderKey!)
            .Select(group => new PdfRedactionTextObjectScope(group.Key, group.ToArray(), reviewedAreas))
            .ToArray();

    private static void AppendPageRenderingResourceIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page) {
        PdfDictionary? current = page.PageDictionary;
        var visited = new HashSet<int>();
        while (current != null) {
            if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                PdfObjectLookup.TryResolveReferenceChain(document.Objects, resourcesObject, out PdfObject? resolvedResources) &&
                resolvedResources is PdfDictionary resources) {
                identity.Append("|R:Font:");
                PdfRedactionImageIdentity.AppendObjectGraph(identity, resources.Items.TryGetValue("Font", out PdfObject? fonts) ? fonts : null, document.Objects);
                identity.Append("|R:ExtGState:");
                PdfRedactionImageIdentity.AppendObjectGraph(identity, resources.Items.TryGetValue("ExtGState", out PdfObject? states) ? states : null, document.Objects);
                identity.Append("|R:Properties:");
                PdfRedactionImageIdentity.AppendObjectGraph(identity, resources.Items.TryGetValue("Properties", out PdfObject? properties) ? properties : null, document.Objects);
                AppendFormRenderingResourceIdentity(identity, document.Objects, resources);
                break;
            }
            if (!current.Items.TryGetValue("Parent", out PdfObject? parentObject) ||
                parentObject is not PdfReference parent ||
                !visited.Add(parent.ObjectNumber) ||
                !PdfObjectLookup.TryGet(document.Objects, parent, out PdfIndirectObject? parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDictionary) return;
            current = parentDictionary;
        }
    }

    private static void AppendFormRenderingResourceIdentity(
        System.Text.StringBuilder identity,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageResources) {
        const int maximumDepth = 64;
        const int maximumContexts = 16384;
        var visited = new HashSet<(PdfStream Form, PdfDictionary Resources)>();
        var formResourceIdentities = new HashSet<string>(StringComparer.Ordinal);
        int contextCount = 0;

        void AppendForms(PdfDictionary resources, int depth) {
            if (depth > maximumDepth) {
                formResourceIdentities.Add(":depth-limit");
                return;
            }
            if (!resources.Items.TryGetValue("XObject", out PdfObject? xObjectValue) ||
                PdfObjectLookup.ResolveChain(objects, xObjectValue) is not PdfDictionary xObjects) return;

            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (PdfObjectLookup.ResolveChain(objects, entry.Value) is not PdfStream form ||
                    PdfObjectLookup.ResolveChain(objects, form.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtype) ? subtype : null) is not PdfName { Name: "Form" }) continue;

                PdfDictionary? declaredResources = form.Dictionary.Items.TryGetValue("Resources", out PdfObject? formResourceValue)
                    ? PdfObjectLookup.ResolveChain(objects, formResourceValue) as PdfDictionary
                    : null;
                PdfDictionary effectiveResources = declaredResources ?? resources;

                var formIdentity = new System.Text.StringBuilder();
                formIdentity.Append(":OC:");
                PdfRedactionImageIdentity.AppendObjectGraph(
                    formIdentity,
                    form.Dictionary.Items.TryGetValue("OC", out PdfObject? optionalContent) ? optionalContent : null,
                    objects);
                formIdentity.Append(":Properties:");
                PdfRedactionImageIdentity.AppendObjectGraph(
                    formIdentity,
                    effectiveResources.Items.TryGetValue("Properties", out PdfObject? properties) ? properties : null,
                    objects);
                if (declaredResources != null) {
                    formIdentity.Append(":Font:");
                    PdfRedactionImageIdentity.AppendObjectGraph(formIdentity, effectiveResources.Items.TryGetValue("Font", out PdfObject? fonts) ? fonts : null, objects);
                    formIdentity.Append(":ExtGState:");
                    PdfRedactionImageIdentity.AppendObjectGraph(formIdentity, effectiveResources.Items.TryGetValue("ExtGState", out PdfObject? states) ? states : null, objects);
                } else {
                    formIdentity.Append(":inherited");
                }
                formResourceIdentities.Add(formIdentity.ToString());

                if (!visited.Add((form, effectiveResources))) continue;
                if (++contextCount > maximumContexts) {
                    formResourceIdentities.Add(":context-limit");
                    return;
                }
                AppendForms(effectiveResources, depth + 1);
            }
        }

        AppendForms(pageResources, 0);
        foreach (string formResourceIdentity in formResourceIdentities.OrderBy(static value => value, StringComparer.Ordinal)) {
            identity.Append("|R:Form").Append(formResourceIdentity);
        }
    }

    private static void AppendUnredactedPathIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas,
        IReadOnlyList<PdfPageDrawingEffectTransition> drawingEffects) {
        IReadOnlyList<PdfPageVisualPrimitive> primitives = page.GetIdentityVisualPrimitives();
        PdfVisualBounds[] visualAreas = pageAreas
            .Select(area => page.TransformBoundsToVisual(area.X, area.Y, area.Right, area.Top))
            .ToArray();
        for (int i = 0; i < primitives.Count; i++) {
            PdfPageVisualPrimitive primitive = primitives[i];
            double strokePadding = primitive.HasStrokePaint ? Math.Max(0D, primitive.StrokeWidth) / 2D : 0D;
            double x = primitive.X - strokePadding;
            double y = primitive.Y - strokePadding;
            double width = primitive.Width + strokePadding * 2D;
            double height = primitive.Height + strokePadding * 2D;
            if (IntersectsReviewedArea(visualAreas, x, y, width, height)) continue;

            identity.Append("|P:").Append((int)primitive.Kind)
                .Append(':').Append(FormatIdentityNumber(primitive.X))
                .Append(',').Append(FormatIdentityNumber(primitive.Y))
                .Append(',').Append(FormatIdentityNumber(primitive.Width))
                .Append(',').Append(FormatIdentityNumber(primitive.Height))
                .Append(',').Append(FormatIdentityNumber(primitive.X1))
                .Append(',').Append(FormatIdentityNumber(primitive.Y1))
                .Append(',').Append(FormatIdentityNumber(primitive.X2))
                .Append(',').Append(FormatIdentityNumber(primitive.Y2))
                .Append(':').Append(FormatIdentityNumber(primitive.StrokeWidth))
                .Append(':').Append((int)primitive.StrokeDashStyle)
                .Append(':').Append(primitive.StrokeLineCap.HasValue ? ((int)primitive.StrokeLineCap.Value).ToString(System.Globalization.CultureInfo.InvariantCulture) : "null")
                .Append(':').Append(primitive.StrokeLineJoin.HasValue ? ((int)primitive.StrokeLineJoin.Value).ToString(System.Globalization.CultureInfo.InvariantCulture) : "null")
                .Append(':').Append((int)primitive.FillRule)
                .Append(':').Append(primitive.FillOpacity.HasValue ? FormatIdentityNumber(primitive.FillOpacity.Value) : "null")
                .Append(':').Append(primitive.StrokeOpacity.HasValue ? FormatIdentityNumber(primitive.StrokeOpacity.Value) : "null");
            AppendStrokeDashIdentity(identity, primitive.StrokeDashPattern);
            AppendIdentityColor(identity, primitive.FillColor);
            AppendIdentityColor(identity, primitive.StrokeColor);
            AppendIdentityGradient(identity, primitive.FillGradient);
            AppendIdentityGradient(identity, primitive.StrokeGradient);
            AppendIdentityGradient(identity, primitive.FillRadialGradient);
            AppendIdentityGradient(identity, primitive.StrokeRadialGradient);
            PdfRedactionImageIdentity.AppendClip(identity, primitive.ClipPath);
            AppendIdentityTilingPattern(identity, primitive.FillTilingPattern);
            AppendIdentityTilingPattern(identity, primitive.StrokeTilingPattern);
            identity.Append(':').Append(primitive.PathCommands.Count);
            for (int commandIndex = 0; commandIndex < primitive.PathCommands.Count; commandIndex++) {
                OfficeIMO.Drawing.OfficePathCommand command = primitive.PathCommands[commandIndex];
                identity.Append(';').Append((int)command.Kind)
                    .Append(',').Append(FormatIdentityNumber(command.Point.X))
                    .Append(',').Append(FormatIdentityNumber(command.Point.Y))
                    .Append(',').Append(FormatIdentityNumber(command.ControlPoint1.X))
                    .Append(',').Append(FormatIdentityNumber(command.ControlPoint1.Y))
                    .Append(',').Append(FormatIdentityNumber(command.ControlPoint2.X))
                    .Append(',').Append(FormatIdentityNumber(command.ControlPoint2.Y));
            }
            AppendDrawingEffectIdentity(identity, document, PdfReadPage.ResolveDrawingEffect(drawingEffects, primitive.PaintOrder, contentOrderKey: primitive.ContentOrderKey));
        }
    }

    private static void AppendStrokeDashIdentity(
        System.Text.StringBuilder identity,
        PdfStrokeDashPattern? pattern) {
        if (!pattern.HasValue) {
            identity.Append(":dash:null");
            return;
        }

        PdfStrokeDashPattern value = pattern.Value;
        identity.Append(":dash:").Append(FormatIdentityNumber(value.Phase));
        AppendIdentityNumbers(identity, value.Array);
    }

    private static void AppendDrawingEffectIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfPageDrawingEffect effect) {
        identity.Append(":effect:")
            .Append((int)effect.BlendMode)
            .Append(',').Append(effect.HasBlendMode ? '1' : '0')
            .Append(',').Append(effect.HasSoftMask ? '1' : '0')
            .Append(',').Append((int)effect.RenderingIntent)
            .Append(',').Append(effect.HasRenderingIntent ? '1' : '0');
        if (effect.SoftMask == null) {
            identity.Append(":smask:null");
            return;
        }
        PdfPageSoftMaskResource softMask = effect.SoftMask;
        identity.Append(":smask:")
            .Append((int)softMask.Mode)
            .Append(',').Append(softMask.HasByteExactBackdrop ? '1' : '0')
            .Append(',').Append(softMask.IsIsolated ? '1' : '0')
            .Append(',').Append(softMask.HasExplicitGroupColorSpace ? '1' : '0');
        AppendIdentityColor(identity, softMask.BackdropColor);
        if (effect.SoftMaskTransform.HasValue) {
            Matrix2D transform = effect.SoftMaskTransform.Value;
            identity.Append(":matrix:").Append(FormatIdentityNumber(transform.A))
                .Append(',').Append(FormatIdentityNumber(transform.B))
                .Append(',').Append(FormatIdentityNumber(transform.C))
                .Append(',').Append(FormatIdentityNumber(transform.D))
                .Append(',').Append(FormatIdentityNumber(transform.E))
                .Append(',').Append(FormatIdentityNumber(transform.F));
        } else {
            identity.Append(":matrix:null");
        }
        identity.Append(":group:");
        PdfRedactionImageIdentity.AppendObjectGraph(identity, softMask.Group, document.Objects);
        identity.Append(":resources:");
        PdfRedactionImageIdentity.AppendObjectGraph(identity, softMask.ParentResources, document.Objects);
    }

    private static void AppendIdentityTilingPattern(
        System.Text.StringBuilder identity,
        PdfPageTilingPatternPaint? pattern) {
        if (pattern is null) {
            identity.Append(":pattern:null");
            return;
        }

        OfficeIMO.Drawing.OfficeTransform transform = pattern.Transform;
        PdfPageTilingPatternResource resource = pattern.Resource;
        identity.Append(":pattern:")
            .Append(FormatIdentityNumber(pattern.Opacity)).Append(':')
            .Append(FormatIdentityNumber(transform.M11)).Append(',')
            .Append(FormatIdentityNumber(transform.M12)).Append(',')
            .Append(FormatIdentityNumber(transform.M21)).Append(',')
            .Append(FormatIdentityNumber(transform.M22)).Append(',')
            .Append(FormatIdentityNumber(transform.OffsetX)).Append(',')
            .Append(FormatIdentityNumber(transform.OffsetY)).Append(':')
            .Append(FormatIdentityNumber(resource.HorizontalStep)).Append(',')
            .Append(FormatIdentityNumber(resource.VerticalStep)).Append(',')
            .Append(FormatIdentityNumber(resource.BoundingBoxX)).Append(',')
            .Append(FormatIdentityNumber(resource.BoundingBoxTop)).Append(':')
            .Append(FormatIdentityNumber(resource.Matrix.A)).Append(',')
            .Append(FormatIdentityNumber(resource.Matrix.B)).Append(',')
            .Append(FormatIdentityNumber(resource.Matrix.C)).Append(',')
            .Append(FormatIdentityNumber(resource.Matrix.D)).Append(',')
            .Append(FormatIdentityNumber(resource.Matrix.E)).Append(',')
            .Append(FormatIdentityNumber(resource.Matrix.F)).Append(':')
            .Append(resource.Uncolored ? '1' : '0')
            .Append(resource.ConsumesInheritedLineState ? '1' : '0')
            .Append(resource.HasMalformedStrictInvocation ? '1' : '0');
        AppendIdentityColor(identity, pattern.Tint);
        identity.Append(':').Append(resource.SourceIdentity);
    }

    private static bool IntersectsReviewedArea(
        PdfVisualBounds[] areas,
        double x,
        double y,
        double width,
        double height) {
        for (int i = 0; i < areas.Length; i++) {
            PdfVisualBounds area = areas[i];
            if (x < area.Right && x + width > area.Left &&
                y < area.Bottom && y + height > area.Top) {
                return true;
            }
        }
        return false;
    }

    private static void AppendIdentityColor(System.Text.StringBuilder identity, OfficeIMO.Drawing.OfficeColor? color) {
        if (!color.HasValue) {
            identity.Append(":null");
            return;
        }
        OfficeIMO.Drawing.OfficeColor value = color.Value;
        identity.Append(':').Append(value.R).Append(',').Append(value.G).Append(',').Append(value.B).Append(',').Append(value.A);
    }

    private static void AppendIdentityPdfColor(System.Text.StringBuilder identity, PdfColor? color) {
        if (!color.HasValue) {
            identity.Append(":null");
            return;
        }
        PdfColor value = color.Value;
        identity.Append(':').Append(FormatIdentityNumber(value.R))
            .Append(',').Append(FormatIdentityNumber(value.G))
            .Append(',').Append(FormatIdentityNumber(value.B));
    }

    private static void AppendIdentityGradient(System.Text.StringBuilder identity, OfficeIMO.Drawing.OfficeLinearGradient? gradient) {
        if (gradient == null) {
            identity.Append(":null");
            return;
        }
        identity.Append(":L,").Append(FormatIdentityNumber(gradient.StartX)).Append(',').Append(FormatIdentityNumber(gradient.StartY))
            .Append(',').Append(FormatIdentityNumber(gradient.EndX)).Append(',').Append(FormatIdentityNumber(gradient.EndY));
        AppendIdentityGradientStops(identity, gradient.Stops);
    }

    private static void AppendIdentityGradient(System.Text.StringBuilder identity, OfficeIMO.Drawing.OfficeRadialGradient? gradient) {
        if (gradient == null) {
            identity.Append(":null");
            return;
        }
        identity.Append(":R,").Append(FormatIdentityNumber(gradient.StartX)).Append(',').Append(FormatIdentityNumber(gradient.StartY))
            .Append(',').Append(FormatIdentityNumber(gradient.StartRadiusX)).Append(',').Append(FormatIdentityNumber(gradient.StartRadiusY))
            .Append(',').Append(FormatIdentityNumber(gradient.EndX)).Append(',').Append(FormatIdentityNumber(gradient.EndY))
            .Append(',').Append(FormatIdentityNumber(gradient.EndRadiusX)).Append(',').Append(FormatIdentityNumber(gradient.EndRadiusY));
        AppendIdentityGradientStops(identity, gradient.Stops);
    }

    private static void AppendIdentityGradientStops(System.Text.StringBuilder identity, IReadOnlyList<OfficeIMO.Drawing.OfficeGradientStop> stops) {
        identity.Append(',').Append(stops.Count);
        for (int i = 0; i < stops.Count; i++) {
            identity.Append(';').Append(FormatIdentityNumber(stops[i].Offset));
            AppendIdentityColor(identity, stops[i].Color);
        }
    }

    private static void AppendUnredactedImageIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page,
        int pageNumber,
        IReadOnlyList<PdfRedactionArea> pageAreas,
        IReadOnlyList<PdfPageDrawingEffectTransition> drawingEffects) {
        IReadOnlyList<PdfImagePlacement> placements = page.GetImagePlacements();
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            if (IntersectsReviewedArea(pageAreas, placement.X, placement.Y, placement.Width, placement.Height)) continue;
            PdfStream? imageStream = null;
            if (placement.ObjectNumber > 0 &&
                document.Objects.TryGetValue(placement.ObjectNumber, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfStream stream) {
                imageStream = stream;
            } else if (placement.InlineImageStream is PdfStream inlineStream) {
                imageStream = inlineStream;
            }

            identity.Append("|I:")
                .Append(pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .Append(':').Append(FormatIdentityNumber(placement.A))
                .Append(',').Append(FormatIdentityNumber(placement.B))
                .Append(',').Append(FormatIdentityNumber(placement.C))
                .Append(',').Append(FormatIdentityNumber(placement.D))
                .Append(',').Append(FormatIdentityNumber(placement.E))
                .Append(',').Append(FormatIdentityNumber(placement.F));
            PdfRedactionImageIdentity.Append(identity, placement, imageStream, document.Objects);
            AppendDrawingEffectIdentity(identity, document, PdfReadPage.ResolveDrawingEffect(drawingEffects, placement.PaintOrder, contentOrderKey: placement.ContentOrderKey));
        }
    }

    private static void AppendUnredactedAnnotationIdentity(
        System.Text.StringBuilder identity,
        PdfReadDocument document,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas,
        IReadOnlyDictionary<int, string> stablePageReferences) {
        IReadOnlyList<PdfAnnotation> annotations = page.GetAnnotationsForContentSafety();
        for (int i = 0; i < annotations.Count; i++) {
            PdfAnnotation annotation = annotations[i];
            if (annotation.HasReadableRectangle &&
                IntersectsReviewedArea(pageAreas, annotation.X1, annotation.Y1, annotation.Width, annotation.Height)) continue;

            identity.Append("|A:");
            AppendIdentityString(identity, annotation.Subtype);
            identity.Append(':').Append(FormatIdentityNumber(annotation.X1))
                .Append(',').Append(FormatIdentityNumber(annotation.Y1))
                .Append(',').Append(FormatIdentityNumber(annotation.X2))
                .Append(',').Append(FormatIdentityNumber(annotation.Y2));
            AppendIdentityString(identity, annotation.Contents);
            AppendIdentityString(identity, annotation.Name);
            AppendIdentityString(identity, annotation.Title);
            AppendIdentityString(identity, annotation.ActionType);
            identity.Append(':').Append(annotation.Flags?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(annotation.HasNormalAppearance ? '1' : '0');
            AppendIdentityString(identity, annotation.AppearanceState);
            AppendIdentityString(identity, annotation.DefaultAppearance);
            AppendIdentityString(identity, annotation.DefaultStyle);
            AppendIdentityString(identity, annotation.RichContents);
            AppendIdentityString(identity, annotation.RichContentsPlainText);
            AppendIdentityNullableNumber(identity, annotation.EffectiveFontSize);
            AppendIdentityPdfColor(identity, annotation.EffectiveTextColor);
            identity.Append(':').Append(annotation.EffectiveTextAlign.HasValue
                ? ((int)annotation.EffectiveTextAlign.Value).ToString(System.Globalization.CultureInfo.InvariantCulture)
                : "null");
            AppendIdentityNullableNumber(identity, annotation.Opacity);
            AppendIdentityNullableNumber(identity, annotation.BorderWidth);
            AppendIdentityString(identity, annotation.BorderStyle);
            AppendIdentityNumbers(identity, annotation.BorderDashPattern);
            AppendIdentityString(identity, annotation.BorderEffectStyle);
            AppendIdentityNullableNumber(identity, annotation.BorderEffectIntensity);
            AppendIdentityNumbers(identity, annotation.RectangleDifferences);
            AppendIdentityNumbers(identity, annotation.CalloutLine);
            AppendIdentityString(identity, annotation.CalloutLineEnding);
            AppendIdentityString(identity, annotation.LineStartEnding);
            AppendIdentityString(identity, annotation.LineEndEnding);
            identity.Append(":appearance:");
            PdfRedactionImageIdentity.AppendObjectGraph(identity, annotation.NormalAppearanceObject, document.Objects);
            AppendIdentityNumbers(identity, annotation.Color);
            AppendIdentityNumbers(identity, annotation.InteriorColor);
            AppendIdentityNumbers(identity, annotation.QuadPoints);
            AppendIdentityNumbers(identity, annotation.LineCoordinates);
            AppendIdentityNumbers(identity, annotation.Vertices);
            for (int pathIndex = 0; pathIndex < annotation.InkList.Count; pathIndex++) {
                AppendIdentityNumbers(identity, annotation.InkList[pathIndex]);
            }
            for (int actionIndex = 0; actionIndex < annotation.AdditionalActions.Count; actionIndex++) {
                PdfAnnotationAdditionalAction action = annotation.AdditionalActions[actionIndex];
                AppendIdentityString(identity, action.TriggerName);
                AppendIdentityString(identity, action.ActionType);
            }
            for (int actionIndex = 0; actionIndex < annotation.ChainedActions.Count; actionIndex++) {
                PdfAnnotationChainedAction action = annotation.ChainedActions[actionIndex];
                AppendIdentityString(identity, action.SourceName);
                AppendIdentityString(identity, action.ActionPath);
                AppendIdentityString(identity, action.ActionType);
            }
            if (annotation.Review != null) {
                AppendIdentityString(identity, annotation.Review.ReplyType);
                AppendIdentityString(identity, annotation.Review.State);
                AppendIdentityString(identity, annotation.Review.StateModel);
                AppendIdentityString(identity, annotation.Review.Subject);
                AppendIdentityString(identity, annotation.Review.Intent);
            }
            if (annotation.SourceDictionary != null) {
                identity.Append(":dictionary:")
                    .Append(PdfRedactionAnnotationIdentity.Compute(
                        annotation.SourceDictionary,
                        document.Objects,
                        stablePageReferences));
            }
        }
    }

    private static void AppendUnredactedLinkIdentity(
        System.Text.StringBuilder identity,
        PdfReadPage page,
        IReadOnlyList<PdfRedactionArea> pageAreas) {
        IReadOnlyList<PdfLinkAnnotation> links = page.GetLinkAnnotations();
        for (int i = 0; i < links.Count; i++) {
            PdfLinkAnnotation link = links[i];
            if (IntersectsReviewedArea(pageAreas, link.X1, link.Y1, link.Width, link.Height)) continue;

            identity.Append("|L:")
                .Append(FormatIdentityNumber(link.X1)).Append(',')
                .Append(FormatIdentityNumber(link.Y1)).Append(',')
                .Append(FormatIdentityNumber(link.X2)).Append(',')
                .Append(FormatIdentityNumber(link.Y2));
            AppendIdentityString(identity, link.Contents);
            AppendIdentityString(identity, link.Uri);
            AppendIdentityString(identity, link.DestinationName);
            AppendIdentityString(identity, link.NamedAction);
            AppendIdentityString(identity, link.RemoteFile);
            AppendIdentityString(identity, link.RemoteDestinationName);
            identity.Append(':').Append(link.DestinationPageNumber?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(link.DestinationMode?.ToString() ?? "null")
                .Append(':').Append(link.RemoteDestinationPageNumber?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "null")
                .Append(':').Append(link.RemoteDestinationMode?.ToString() ?? "null");
            AppendIdentityNullableNumber(identity, link.DestinationLeft);
            AppendIdentityNullableNumber(identity, link.DestinationTop);
            AppendIdentityNullableNumber(identity, link.DestinationBottom);
            AppendIdentityNullableNumber(identity, link.DestinationRight);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationLeft);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationTop);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationBottom);
            AppendIdentityNullableNumber(identity, link.RemoteDestinationRight);
        }
    }

    private static void AppendIdentityNullableNumber(System.Text.StringBuilder identity, double? value) =>
        identity.Append(':').Append(value.HasValue ? FormatIdentityNumber(value.Value) : "null");

    private static void AppendIdentityString(System.Text.StringBuilder identity, string? value) {
        if (value == null) {
            identity.Append(":null");
            return;
        }
        identity.Append(':')
            .Append(value.Length.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append(':').Append(value);
    }

    private static void AppendIdentityNumbers(
        System.Text.StringBuilder identity,
        IReadOnlyList<double> values) {
        identity.Append(':').Append(values.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        for (int i = 0; i < values.Count; i++) {
            identity.Append(',').Append(FormatIdentityNumber(values[i]));
        }
    }

    private static bool IntersectsReviewedArea(
        IReadOnlyList<PdfRedactionArea> areas,
        double x,
        double y,
        double width,
        double height) {
        for (int i = 0; i < areas.Count; i++) {
            PdfRedactionArea area = areas[i];
            if (x < area.X + area.Width && x + width > area.X &&
                y < area.Y + area.Height && y + height > area.Y) {
                return true;
            }
        }
        return false;
    }

    private static string FormatIdentityNumber(double value) =>
        value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

    private static string ComputeIdentityHash(string value) =>
        ComputeIdentityHash(System.Text.Encoding.UTF8.GetBytes(value));

    private static string ComputeIdentityHash(byte[] value) {
#if NET6_0_OR_GREATER
        return Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(value));
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(value));
#endif
    }

    private static string FormatPageBoxIdentity(PdfPageBox? box) {
        if (box == null) {
            return "null";
        }

        return string.Join(",", new[] {
            box.Left.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            box.Bottom.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            box.Right.ToString("R", System.Globalization.CultureInfo.InvariantCulture),
            box.Top.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
        });
    }
}
