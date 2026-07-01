using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared fluent image-export surface for document-level batch exports.
/// </summary>
/// <typeparam name="TBuilder">Concrete builder type returned from fluent methods.</typeparam>
/// <typeparam name="TOptions">Document-specific image export options.</typeparam>
public abstract class OfficeImageExportBatchBuilder<TBuilder, TOptions>
    where TBuilder : OfficeImageExportBatchBuilder<TBuilder, TOptions>
    where TOptions : OfficeImageExportOptions {
    private readonly Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> _export;
    private OfficeImageExportFormat _format = OfficeImageExportFormat.Png;

    /// <summary>
    /// Creates a fluent batch export builder over an existing document-specific export function.
    /// </summary>
    protected OfficeImageExportBatchBuilder(TOptions options, Func<OfficeImageExportFormat, TOptions, IReadOnlyList<OfficeImageExportResult>> export) {
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

    /// <summary>Exports all selected images using the currently configured format and options.</summary>
    public IReadOnlyList<OfficeImageExportResult> Export() => _export(_format, Options);

    /// <summary>Exports selected images as PNG using the currently configured options.</summary>
    public IReadOnlyList<OfficeImageExportResult> ExportPng() {
        _format = OfficeImageExportFormat.Png;
        return Export();
    }

    /// <summary>Exports selected images as SVG using the currently configured options.</summary>
    public IReadOnlyList<OfficeImageExportResult> ExportSvg() {
        _format = OfficeImageExportFormat.Svg;
        return Export();
    }

    /// <summary>Saves all selected images to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> Save(string folderPath) {
        if (string.IsNullOrWhiteSpace(folderPath)) {
            throw new ArgumentException("Output folder cannot be null or whitespace.", nameof(folderPath));
        }

        string fullFolder = Path.GetFullPath(folderPath);
        Directory.CreateDirectory(fullFolder);
        IReadOnlyList<OfficeImageExportResult> results = Export();
        string extension = _format == OfficeImageExportFormat.Svg ? ".svg" : ".png";
        var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < results.Count; i++) {
            OfficeImageExportResult result = results[i];
            string name = string.IsNullOrWhiteSpace(result.Name)
                ? "image-" + (i + 1).ToString(CultureInfo.InvariantCulture)
                : result.Name!;
            string fileName = GetUniqueFileName(SanitizeFileName(name), extension, usedFileNames);
            File.WriteAllBytes(Path.Combine(fullFolder, fileName), result.Bytes);
        }

        return results;
    }

    /// <summary>Saves all selected images to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> SaveTo(string folderPath) => Save(folderPath);

    /// <summary>Saves selected images as PNG files to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> SaveAsPng(string folderPath) {
        _format = OfficeImageExportFormat.Png;
        return Save(folderPath);
    }

    /// <summary>Saves selected images as PNG files to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> SavePng(string folderPath) => SaveAsPng(folderPath);

    /// <summary>Saves selected images as SVG files to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> SaveAsSvg(string folderPath) {
        _format = OfficeImageExportFormat.Svg;
        return Save(folderPath);
    }

    /// <summary>Saves selected images as SVG files to a folder.</summary>
    public IReadOnlyList<OfficeImageExportResult> SaveSvg(string folderPath) => SaveAsSvg(folderPath);

    private static string SanitizeFileName(string name) {
        char[] invalid = Path.GetInvalidFileNameChars();
        char[] chars = name.ToCharArray();
        for (int i = 0; i < chars.Length; i++) {
            if (Array.IndexOf(invalid, chars[i]) >= 0) {
                chars[i] = '_';
            }
        }

        return new string(chars).Trim();
    }

    private static string GetUniqueFileName(string baseName, string extension, ISet<string> usedFileNames) {
        if (string.IsNullOrWhiteSpace(baseName)) {
            baseName = "image";
        }

        string candidate = baseName + extension;
        if (usedFileNames.Add(candidate)) {
            return candidate;
        }

        int suffix = 2;
        do {
            candidate = baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture) + extension;
            suffix++;
        } while (!usedFileNames.Add(candidate));

        return candidate;
    }

    private TBuilder This => (TBuilder)this;
}
