using System.Diagnostics;

namespace OfficeIMO.Drawing;

/// <summary>Describes one native chart for <see cref="OfficeChartImageExporter"/>.</summary>
internal sealed class OfficeChartImageSource {
    private readonly Func<double, double, List<OfficeImageExportDiagnostic>, OfficeChartSnapshot> _createSnapshot;

    /// <param name="producer">Format family used in conversion reports, such as "Word".</param>
    /// <param name="name">Chart name reported on the result.</param>
    /// <param name="source">Location reported on diagnostics.</param>
    /// <param name="frameWidth">Chart width in its document.</param>
    /// <param name="frameHeight">Chart height in its document.</param>
    /// <param name="createSnapshot">Projects the chart at a logical extent; throws <see cref="NotSupportedException"/> with the reason when it cannot.</param>
    internal OfficeChartImageSource(string producer, string? name, string source, OfficeLength frameWidth, OfficeLength frameHeight,
        Func<double, double, List<OfficeImageExportDiagnostic>, OfficeChartSnapshot> createSnapshot) {
        Producer = producer;
        Name = name;
        Source = source;
        FrameWidth = frameWidth;
        FrameHeight = frameHeight;
        _createSnapshot = createSnapshot ?? throw new ArgumentNullException(nameof(createSnapshot));
    }

    internal string Producer { get; }
    internal string? Name { get; }
    internal string Source { get; }
    internal OfficeLength FrameWidth { get; }
    internal OfficeLength FrameHeight { get; }

    internal OfficeChartSnapshot CreateSnapshot(double width, double height, List<OfficeImageExportDiagnostic> diagnostics) =>
        _createSnapshot(width, height, diagnostics);
}

/// <summary>Renders one shared chart snapshot as a standalone image for the native chart owners.</summary>
internal static class OfficeChartImageExporter {
    internal const string UnsupportedChartMessage = "The chart cannot be rendered by the shared Office chart renderer: ";

    internal static OfficeImageExportResult ToImage(OfficeChartImageSource source, OfficeChartImageExportOptions? options,
        CancellationToken cancellationToken) {
        OfficeChartImageExportOptions resolved = Normalize(options);
        return CreateOperation(source, resolved).Export(cancellationToken);
    }

    internal static OfficeConversionReport SaveAsImage(OfficeChartImageSource source, string path,
        OfficeChartImageExportOptions? options, CancellationToken cancellationToken) {
        var timer = Stopwatch.StartNew();
        OfficeChartImageExportOptions resolved = Normalize(options);
        OfficeImageExportResult result = CreateOperation(source, resolved).Save(path, cancellationToken);
        return OfficeImageExportEvidence.CreateReport(source.Producer, result, timer.Elapsed);
    }

    internal static OfficeConversionReport SaveAsImage(OfficeChartImageSource source, Stream stream,
        OfficeChartImageExportOptions? options, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(stream);
        var timer = Stopwatch.StartNew();
        OfficeChartImageExportOptions resolved = Normalize(options);
        OfficeImageExportResult result = CreateOperation(source, resolved).Save(stream, cancellationToken);
        return OfficeImageExportEvidence.CreateReport(source.Producer, result, timer.Elapsed);
    }

    /// <summary>Renders an existing chart snapshot; used by owners that keep their own image-export options.</summary>
    internal static OfficeImageExportResult Export(OfficeChartSnapshot snapshot, OfficeImageExportFormat format,
        OfficeImageExportOptions options, List<OfficeImageExportDiagnostic> diagnostics, string? name, string source,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot, useMinimumCanvas: false, diagnostics);
        drawing.Fonts.AddRange(options.Fonts);
        drawing.AppendFontDiagnostics(diagnostics, source);

        if (format == OfficeImageExportFormat.Svg) {
            var svgDrawing = new OfficeDrawing(drawing.Width, drawing.Height);
            OfficeShape background = OfficeShape.Rectangle(drawing.Width, drawing.Height);
            background.FillColor = options.BackgroundColor;
            background.StrokeWidth = 0D;
            svgDrawing.AddShape(background, 0D, 0D);
            svgDrawing.AddDrawing(drawing, 0D, 0D);
            double scale = options.GetEffectiveScale(svgDrawing.Width, svgDrawing.Height);
            byte[] svgBytes = OfficeDrawingSvgExporter.ToSvgBytes(svgDrawing, scale, OfficeSvgSizeUnit.Pixel,
                imageCodec: null, resourceIdPrefix: null, maximumUtf8Bytes: options.MaximumTotalEncodedBytes,
                cancellationToken);
            if (!OfficeImageReader.TryIdentifyByContent(svgBytes, ".svg", out OfficeImageInfo svgInfo)) {
                throw new InvalidDataException("The rendered chart SVG dimensions could not be identified.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            return options.EnsureAccepted(new OfficeImageExportResult(format, svgInfo.Width, svgInfo.Height, svgBytes,
                name, source, diagnostics.AsReadOnly()));
        }

        if (!format.IsRaster()) {
            throw new ArgumentOutOfRangeException(nameof(format), "The selected image format is not supported.");
        }

        OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(drawing.Width, drawing.Height, format, options, source);
        if (plan.Diagnostic != null) diagnostics.Add(plan.Diagnostic);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            Scale = plan.Limit.Scale,
            Background = options.BackgroundColor,
            TextShapingProvider = options.TextShapingProvider,
            TextShapingLanguage = options.TextShapingLanguage,
            DiagnosticSink = diagnostics,
            DiagnosticSource = source,
            CancellationToken = cancellationToken
        });
        byte[] bytes = OfficeRasterImageEncoder.Encode(image, format, plan.CreateEncodingOptions(),
            options.MaximumTotalEncodedBytes, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return options.EnsureAccepted(new OfficeImageExportResult(format, image.Width, image.Height, bytes, name, source,
            diagnostics.AsReadOnly()));
    }

    private static OfficeChartImageExportOptions Normalize(OfficeChartImageExportOptions? options) {
        OfficeChartImageExportOptions resolved = options?.Clone() ?? new OfficeChartImageExportOptions();
        resolved.Validate();
        return resolved;
    }

    private static OfficeImageExportOperation<OfficeChartImageExportOptions> CreateOperation(OfficeChartImageSource source,
        OfficeChartImageExportOptions options) {
        ArgumentNullException.ThrowIfNull(source);
        return new OfficeImageExportOperation<OfficeChartImageExportOptions>(options, options.Format, options.FileConflictPolicy,
            (format, effective) => Render(source, format, effective, CancellationToken.None),
            (format, effective, token) => Render(source, format, effective, token));
    }

    private static OfficeImageExportResult Render(OfficeChartImageSource source, OfficeImageExportFormat format,
        OfficeChartImageExportOptions options, CancellationToken cancellationToken) =>
        OfficeImageExportExecutionScope.Run(options, cancellationToken, token => {
            double units = options.LogicalUnitsPerInch;
            double width = (options.Width ?? source.FrameWidth).ToInches() * units;
            double height = (options.Height ?? source.FrameHeight).ToInches() * units;
            if (!(width > 0D) || !(height > 0D)) {
                throw new NotSupportedException(UnsupportedChartMessage + "the chart frame has no size.");
            }
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeChartSnapshot snapshot = source.CreateSnapshot(width, height, diagnostics);
            return Export(snapshot, format, options, diagnostics, source.Name, source.Source, token);
        });
}
