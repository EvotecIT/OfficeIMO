using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared fluent image-export surface used by OfficeIMO document packages.
/// </summary>
/// <typeparam name="TBuilder">Concrete builder type returned from fluent methods.</typeparam>
/// <typeparam name="TOptions">Document-specific image export options.</typeparam>
public abstract class OfficeImageExportBuilder<TBuilder, TOptions>
    where TBuilder : OfficeImageExportBuilder<TBuilder, TOptions>
    where TOptions : OfficeImageExportOptions {
    private readonly Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> _export;
    private OfficeImageExportFormat _format = OfficeImageExportFormat.Png;

    /// <summary>
    /// Creates a fluent export builder over an existing document-specific export function.
    /// </summary>
    protected OfficeImageExportBuilder(TOptions options, Func<OfficeImageExportFormat, TOptions, OfficeImageExportResult> export) {
        Options = options ?? throw new ArgumentNullException(nameof(options));
        _export = export ?? throw new ArgumentNullException(nameof(export));
    }

    /// <summary>Document-specific options being configured by this builder.</summary>
    protected TOptions Options { get; }

    /// <summary>Configures PNG output.</summary>
    public TBuilder AsPng() {
        _format = OfficeImageExportFormat.Png;
        return This;
    }

    /// <summary>Configures SVG output.</summary>
    public TBuilder AsSvg() {
        _format = OfficeImageExportFormat.Svg;
        return This;
    }

    /// <summary>Configures the output image format.</summary>
    public TBuilder As(OfficeImageExportFormat format) {
        _format = format;
        return This;
    }

    /// <summary>Sets the output scale.</summary>
    public TBuilder WithScale(double scale) {
        OfficeImageExportOptions.ValidateScale(scale);
        Options.Scale = scale;
        return This;
    }

    /// <summary>Sets the output scale using a compact fluent alias.</summary>
    public TBuilder AtScale(double scale) => WithScale(scale);

    /// <summary>Sets the export background color.</summary>
    public TBuilder WithBackground(OfficeColor color) {
        Options.BackgroundColor = color;
        return This;
    }

    /// <summary>Sets the export background from a named color or hexadecimal color value.</summary>
    public TBuilder WithBackground(string color) => WithBackground(OfficeColor.Parse(color));

    /// <summary>Sets the export background color using a compact fluent alias.</summary>
    public TBuilder OnBackground(OfficeColor color) => WithBackground(color);

    /// <summary>Sets the export background from a named color or hexadecimal color value using a compact fluent alias.</summary>
    public TBuilder OnBackground(string color) => WithBackground(color);

    /// <summary>Uses an opaque white export background.</summary>
    public TBuilder OnWhiteBackground() => WithBackground(OfficeColor.White);

    /// <summary>Uses an opaque white export background.</summary>
    public TBuilder WhiteBackground() => OnWhiteBackground();

    /// <summary>Uses a transparent export background when the selected format supports transparency.</summary>
    public TBuilder OnTransparentBackground() => WithBackground(OfficeColor.Transparent);

    /// <summary>Uses a transparent export background when the selected format supports transparency.</summary>
    public TBuilder TransparentBackground() => OnTransparentBackground();

    /// <summary>Configures a standard preview profile: PNG, 1x scale, white background.</summary>
    public TBuilder ForPreview() {
        _format = OfficeImageExportFormat.Png;
        Options.Scale = 1D;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Configures a standard preview profile: PNG, 1x scale, white background.</summary>
    public TBuilder Preview() => ForPreview();

    /// <summary>Configures a high-resolution profile: PNG, 2x scale, white background.</summary>
    public TBuilder ForHighResolution() {
        _format = OfficeImageExportFormat.Png;
        Options.Scale = 2D;
        Options.BackgroundColor = OfficeColor.White;
        return This;
    }

    /// <summary>Configures a high-resolution profile: PNG, 2x scale, white background.</summary>
    public TBuilder HighResolution() => ForHighResolution();

    /// <summary>Exports using the currently configured format and options.</summary>
    public OfficeImageExportResult Export() => _export(_format, Options);

    /// <summary>Exports as PNG using the currently configured options.</summary>
    public OfficeImageExportResult ExportPng() {
        _format = OfficeImageExportFormat.Png;
        return Export();
    }

    /// <summary>Exports as SVG using the currently configured options.</summary>
    public OfficeImageExportResult ExportSvg() {
        _format = OfficeImageExportFormat.Svg;
        return Export();
    }

    /// <summary>Exports and returns the encoded image bytes.</summary>
    public byte[] ToBytes() => Export().Bytes;

    /// <summary>Exports as PNG and returns the encoded image bytes.</summary>
    public byte[] ToPngBytes() => ExportPng().Bytes;

    /// <summary>Exports as PNG and returns the encoded image bytes.</summary>
    public byte[] ToPng() => ToPngBytes();

    /// <summary>Exports as SVG and returns the encoded SVG bytes.</summary>
    public byte[] ToSvgBytes() => ExportSvg().Bytes;

    /// <summary>Exports as SVG and returns the SVG XML text.</summary>
    public string ToSvgString() {
        OfficeImageExportResult result = _export(OfficeImageExportFormat.Svg, Options);
        return Encoding.UTF8.GetString(result.Bytes);
    }

    /// <summary>Exports as SVG and returns the SVG XML text.</summary>
    public string ToSvg() => ToSvgString();

    /// <summary>Saves the exported image to a file path.</summary>
    public OfficeImageExportResult Save(string path) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        OfficeImageExportResult result = Export();
        string fullPath = Path.GetFullPath(path);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(directory)) {
            Directory.CreateDirectory(directory!);
        }

        File.WriteAllBytes(fullPath, result.Bytes);
        return result;
    }

    /// <summary>Saves the exported image to a file path.</summary>
    public OfficeImageExportResult SaveTo(string path) => Save(path);

    /// <summary>Saves the exported image as PNG to a file path.</summary>
    public OfficeImageExportResult SaveAsPng(string path) {
        _format = OfficeImageExportFormat.Png;
        return Save(path);
    }

    /// <summary>Saves the exported image as PNG to a file path.</summary>
    public OfficeImageExportResult SavePng(string path) => SaveAsPng(path);

    /// <summary>Saves the exported image as SVG to a file path.</summary>
    public OfficeImageExportResult SaveAsSvg(string path) {
        _format = OfficeImageExportFormat.Svg;
        return Save(path);
    }

    /// <summary>Saves the exported image as SVG to a file path.</summary>
    public OfficeImageExportResult SaveSvg(string path) => SaveAsSvg(path);

    /// <summary>Writes the exported image to a stream.</summary>
    public OfficeImageExportResult Save(Stream stream) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        OfficeImageExportResult result = Export();
        stream.Write(result.Bytes, 0, result.Bytes.Length);
        return result;
    }

    /// <summary>Writes the exported image to a stream.</summary>
    public OfficeImageExportResult SaveTo(Stream stream) => Save(stream);

    /// <summary>Writes the exported image as PNG to a stream.</summary>
    public OfficeImageExportResult SaveAsPng(Stream stream) {
        _format = OfficeImageExportFormat.Png;
        return Save(stream);
    }

    /// <summary>Writes the exported image as PNG to a stream.</summary>
    public OfficeImageExportResult SavePng(Stream stream) => SaveAsPng(stream);

    /// <summary>Writes the exported image as SVG to a stream.</summary>
    public OfficeImageExportResult SaveAsSvg(Stream stream) {
        _format = OfficeImageExportFormat.Svg;
        return Save(stream);
    }

    /// <summary>Writes the exported image as SVG to a stream.</summary>
    public OfficeImageExportResult SaveSvg(Stream stream) => SaveAsSvg(stream);

    private TBuilder This => (TBuilder)this;
}
