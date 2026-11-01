using System;
using System.IO;
using System.Management.Automation;
using ChartForgeX.VisualArtifacts;
using OfficeIMO;
using OfficeIMO.ChartForgeX;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.Visuals;

/// <summary>Provides shared, typed ChartForgeX-to-OfficeIMO conversion parameters.</summary>
public abstract class OfficeVisualCommandBase : PSCmdlet
{
    /// <summary>SVG fidelity policy used by OfficeIMO.ChartForgeX.</summary>
    [Parameter(ParameterSetName = "Visual")]
    public OfficeVisualSvgPolicy SvgPolicy { get; set; } = OfficeVisualSvgPolicy.PreserveVector;

    /// <summary>Optional output width; numbers mean points and unit suffixes are accepted.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Optional output height; numbers mean points and unit suffixes are accepted.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Conversion factor from ChartForgeX pixels to Office points.</summary>
    [Parameter(ParameterSetName = "Visual")]
    public double PointsPerPixel { get; set; } = 0.75D;

    /// <summary>Optional OfficeIMO SVG import element limit.</summary>
    [Parameter(ParameterSetName = "Visual")]
    public int? MaximumSvgElements { get; set; }

    /// <summary>Optional maximum SVG viewport width or height. Increase the safe default only for trusted input.</summary>
    [Parameter(ParameterSetName = "Visual")]
    public double? MaximumSvgViewportDimension { get; set; }

    /// <summary>Optional maximum SVG viewport area in pixels. Increase the safe default only for trusted input.</summary>
    [Parameter(ParameterSetName = "Visual")]
    public long? MaximumSvgViewportPixelCount { get; set; }

    /// <summary>Optional stable identifier for an SVG file input.</summary>
    [Parameter(ParameterSetName = "Visual")]
    public string? Id { get; set; }

    /// <summary>Optional title for an SVG file input.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Accessible description exposed by the owning command.</summary>
    protected abstract string? VisualAlternativeText { get; }

    /// <summary>Parameter name used when validating metadata overrides.</summary>
    protected virtual string AlternativeTextParameterName => "AlternativeText";

    /// <summary>Converts an artifact or reuses an existing conversion result.</summary>
    /// <param name="inputObject">Portable source or prepared visual.</param>
    /// <param name="nativePlacement">Leave dimensions and alternative text to the native placement owner without changing reusable conversions.</param>
    protected OfficeVisualConversionResult ResolveVisual(object inputObject, bool nativePlacement = false)
    {
        if (inputObject is PathInfo directPathInfo)
        {
            return ConvertSvgFile(directPathInfo.ProviderPath, nativePlacement);
        }

        PSObject input = PSObject.AsPSObject(inputObject);
        object value = input.BaseObject;
        if (value is OfficeVisualConversionResult converted)
        {
            RejectConversionOverrides(nativePlacement);
            return converted;
        }
        if (value is OfficeVisualSource source)
        {
            RejectSourceMetadataOverrides(nativePlacement);
            return source.ToOfficeVisual(CreateOptions(nativePlacement));
        }

        if (value is FileInfo fileInfo)
        {
            return ConvertSvgFile(fileInfo.FullName, nativePlacement);
        }
        if (value is PathInfo pathInfo)
        {
            return ConvertSvgFile(pathInfo.ProviderPath, nativePlacement);
        }
        if (value is string path)
        {
            return ConvertSvgFile(SessionState.Path.GetUnresolvedProviderPathFromPSPath(path), nativePlacement);
        }

        if (input.TypeNames.Contains("ImagePlayground.VisualArtifact"))
        {
            return ConvertPortableVisual(input, nativePlacement);
        }

        if (value is not VisualArtifact artifact)
        {
            throw new PSArgumentException(
                "InputObject must be a ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path. Received " +
                inputObject.GetType().FullName + ".",
                nameof(inputObject));
        }

        return artifact.ToOfficeVisual(CreateOptions(nativePlacement));
    }

    private OfficeVisualConversionResult ConvertPortableVisual(PSObject input, bool nativePlacement)
    {
        object? payload = input.Properties["OfficeVisualSvg"]?.Value;
        if (payload is not byte[] svgBytes || svgBytes.Length == 0)
        {
            throw new PSArgumentException(
                "ImagePlayground.VisualArtifact input must provide non-empty OfficeVisualSvg bytes.",
                nameof(input));
        }

        var source = new OfficeVisualSource(svgBytes)
        {
            Id = ResolvePortableText(input, "OfficeVisualId", Id),
            Title = ResolvePortableText(input, "OfficeVisualTitle", Title),
            AlternativeText = ResolvePortableText(input, "OfficeVisualAlternativeText", VisualAlternativeText)
        };
        return source.ToOfficeVisual(CreateOptions(nativePlacement));
    }

    private static string ResolvePortableText(PSObject input, string propertyName, string? overrideValue)
    {
        if (!string.IsNullOrWhiteSpace(overrideValue))
        {
            return overrideValue!;
        }

        return input.Properties[propertyName]?.Value as string ?? string.Empty;
    }

    private void RejectConversionOverrides(bool nativePlacement)
    {
        string[] names = { nameof(SvgPolicy), nameof(Width), nameof(Height), nameof(PointsPerPixel), nameof(MaximumSvgElements), nameof(MaximumSvgViewportDimension), nameof(MaximumSvgViewportPixelCount), nameof(Id), nameof(Title), AlternativeTextParameterName };
        foreach (string name in names)
        {
            if (nativePlacement && (name == nameof(Width) || name == nameof(Height) || name == AlternativeTextParameterName || name == nameof(Title))) continue;
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisualConversionResult.", name);
            }
        }
    }

    private void RejectSourceMetadataOverrides(bool nativePlacement)
    {
        foreach (string name in new[] { nameof(Id), nameof(Title), AlternativeTextParameterName })
        {
            if (nativePlacement && (name == AlternativeTextParameterName || name == nameof(Title))) continue;
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisualSource.", name);
            }
        }
    }

    private OfficeVisualConversionResult ConvertSvgFile(string path, bool nativePlacement)
    {
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"SVG visual '{fullPath}' was not found.", fullPath);
        }
        if (!string.Equals(Path.GetExtension(fullPath), ".svg", StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException("Portable Office visual input must be an .svg file.", nameof(path));
        }
        var source = new OfficeVisualSource(File.ReadAllBytes(fullPath))
        {
            Id = string.IsNullOrWhiteSpace(Id) ? Path.GetFileNameWithoutExtension(fullPath) : Id!,
            Title = Title ?? string.Empty,
            AlternativeText = VisualAlternativeText ?? string.Empty
        };
        return source.ToOfficeVisual(CreateOptions(nativePlacement));
    }

    private OfficeVisualConversionOptions CreateOptions(bool nativePlacement)
    {
        var options = new OfficeVisualConversionOptions
        {
            SvgPolicy = SvgPolicy,
            PointsPerPixel = PointsPerPixel,
            WidthPoints = nativePlacement ? null : Width?.ToPoints(),
            HeightPoints = nativePlacement ? null : Height?.ToPoints()
        };
        if (MaximumSvgElements.HasValue)
        {
            options.MaximumSvgElements = MaximumSvgElements.Value;
        }
        if (MaximumSvgViewportDimension.HasValue)
        {
            options.MaximumSvgViewportDimension = MaximumSvgViewportDimension.Value;
        }
        if (MaximumSvgViewportPixelCount.HasValue)
        {
            options.MaximumSvgViewportPixels = MaximumSvgViewportPixelCount.Value;
        }
        return options;
    }
}
