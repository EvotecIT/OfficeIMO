using System;
using System.IO;
using System.Management.Automation;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;

namespace PSWriteOffice.Services.Visuals;

/// <summary>Provides shared, typed ChartForgeX-to-OfficeIMO conversion parameters.</summary>
public abstract class OfficeVisualCommandBase : PSCmdlet
{
    /// <summary>SVG fidelity policy used by OfficeIMO.ChartForgeX.</summary>
    [Parameter]
    public OfficeVisualSvgPolicy SvgPolicy { get; set; } = OfficeVisualSvgPolicy.PreserveVector;

    /// <summary>Optional output width in Office points.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Optional output height in Office points.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Conversion factor from ChartForgeX pixels to Office points.</summary>
    [Parameter]
    public double PointsPerPixel { get; set; } = 0.75D;

    /// <summary>Optional OfficeIMO SVG import element limit.</summary>
    [Parameter]
    public int? MaximumSvgElements { get; set; }

    /// <summary>Optional maximum SVG viewport width or height. Increase the safe default only for trusted input.</summary>
    [Parameter]
    public double? MaximumSvgViewportDimension { get; set; }

    /// <summary>Optional maximum SVG viewport area in pixels. Increase the safe default only for trusted input.</summary>
    [Parameter]
    public double? MaximumSvgViewportPixels { get; set; }

    /// <summary>Optional stable identifier for an SVG file input.</summary>
    [Parameter]
    public string? Id { get; set; }

    /// <summary>Optional title for an SVG file input.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional accessible description for an SVG file input.</summary>
    [Parameter]
    public string? AlternativeText { get; set; }

    /// <summary>Converts an artifact or reuses an existing conversion result.</summary>
    protected OfficeVisualConversionResult ResolveVisual(object inputObject)
    {
        if (inputObject is PathInfo directPathInfo)
        {
            return ConvertSvgFile(directPathInfo.ProviderPath);
        }

        PSObject input = PSObject.AsPSObject(inputObject);
        object value = input.BaseObject;
        if (value is OfficeVisualConversionResult converted)
        {
            RejectConversionOverrides();
            return converted;
        }
        if (value is OfficeVisualSource source)
        {
            RejectSourceMetadataOverrides();
            return source.ToOfficeVisual(CreateOptions());
        }

        if (value is FileInfo fileInfo)
        {
            return ConvertSvgFile(fileInfo.FullName);
        }
        if (value is PathInfo pathInfo)
        {
            return ConvertSvgFile(pathInfo.ProviderPath);
        }
        if (value is string path)
        {
            return ConvertSvgFile(SessionState.Path.GetUnresolvedProviderPathFromPSPath(path));
        }

        if (input.TypeNames.Contains("ImagePlayground.VisualArtifact"))
        {
            return ConvertPortableVisual(input);
        }

        if (value is not VisualArtifact artifact)
        {
            throw new PSArgumentException(
                "InputObject must be a ChartForgeX VisualArtifact, OfficeVisualSource, OfficeVisualConversionResult, or SVG file path. Received " +
                inputObject.GetType().FullName + ".",
                nameof(inputObject));
        }

        return artifact.ToOfficeVisual(CreateOptions());
    }

    private OfficeVisualConversionResult ConvertPortableVisual(PSObject input)
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
            AlternativeText = ResolvePortableText(input, "OfficeVisualAlternativeText", AlternativeText)
        };
        return source.ToOfficeVisual(CreateOptions());
    }

    private static string ResolvePortableText(PSObject input, string propertyName, string? overrideValue)
    {
        if (!string.IsNullOrWhiteSpace(overrideValue))
        {
            return overrideValue!;
        }

        return input.Properties[propertyName]?.Value as string ?? string.Empty;
    }

    private void RejectConversionOverrides()
    {
        string[] names = { nameof(SvgPolicy), nameof(Width), nameof(Height), nameof(PointsPerPixel), nameof(MaximumSvgElements), nameof(MaximumSvgViewportDimension), nameof(MaximumSvgViewportPixels), nameof(Id), nameof(Title), nameof(AlternativeText) };
        foreach (string name in names)
        {
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisualConversionResult.", name);
            }
        }
    }

    private void RejectSourceMetadataOverrides()
    {
        foreach (string name in new[] { nameof(Id), nameof(Title), nameof(AlternativeText) })
        {
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisualSource.", name);
            }
        }
    }

    private OfficeVisualConversionResult ConvertSvgFile(string path)
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
            AlternativeText = AlternativeText ?? string.Empty
        };
        return source.ToOfficeVisual(CreateOptions());
    }

    private OfficeVisualConversionOptions CreateOptions()
    {
        var options = new OfficeVisualConversionOptions
        {
            SvgPolicy = SvgPolicy,
            PointsPerPixel = PointsPerPixel,
            WidthPoints = Width,
            HeightPoints = Height
        };
        if (MaximumSvgElements.HasValue)
        {
            options.MaximumSvgElements = MaximumSvgElements.Value;
        }
        if (MaximumSvgViewportDimension.HasValue)
        {
            options.MaximumSvgViewportDimension = MaximumSvgViewportDimension.Value;
        }
        if (MaximumSvgViewportPixels.HasValue)
        {
            options.MaximumSvgViewportPixels = MaximumSvgViewportPixels.Value;
        }
        return options;
    }
}
