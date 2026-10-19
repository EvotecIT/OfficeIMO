using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;

namespace PSWriteOffice.Services.Visuals;

/// <summary>Provides shared parameters and ALC-safe input handling for native editable Visio projection.</summary>
public abstract class OfficeVisioVisualCommandBase : PSCmdlet
{
    private List<byte>? _pipelineBytes;
    private bool _nonBytePipelineInputSeen;

    /// <summary>Name of the generated Visio page.</summary>
    [Parameter]
    public string PageName { get; set; } = "Visual Artifact";

    /// <summary>Use the CFX natural pixel size as the minimum Visio page size.</summary>
    [Parameter]
    public SwitchParameter UseNaturalPageSize { get; set; }

    /// <summary>Pixel density used with -UseNaturalPageSize.</summary>
    [Parameter]
    public double PixelsPerInch { get; set; } = 96D;

    /// <summary>Do not add the artifact title as an editable Visio title.</summary>
    [Parameter]
    public SwitchParameter NoTitle { get; set; }

    /// <summary>Do not create native Visio containers for CFX groups or lanes.</summary>
    [Parameter]
    public SwitchParameter NoGroups { get; set; }

    /// <summary>Do not copy CFX metadata, ports, and details into Visio Shape Data.</summary>
    [Parameter]
    public SwitchParameter NoShapeData { get; set; }

    /// <summary>Do not copy safe CFX links onto native Visio shapes and connectors.</summary>
    [Parameter]
    public SwitchParameter NoHyperlinks { get; set; }

    /// <summary>
    /// Buffers scalar bytes emitted when PowerShell enumerates a piped <see cref="byte"/> array.
    /// Non-byte inputs remain available to the derived cmdlet's normal per-record behavior.
    /// </summary>
    protected bool BufferPipelineByte(object inputObject)
    {
        object value = PSObject.AsPSObject(inputObject).BaseObject;
        if (value is byte item)
        {
            if (_nonBytePipelineInputSeen)
            {
                throw MixedPipelineInput();
            }

            int bufferedCount = _pipelineBytes?.Count ?? 0;
            if (bufferedCount >= VisualArtifactInterchangeEnvelope.MaximumJsonUtf8Bytes)
            {
                throw new PSArgumentOutOfRangeException(
                    "InputObject",
                    bufferedCount + 1,
                    $"CFX interchange UTF-8 JSON must not exceed {VisualArtifactInterchangeEnvelope.MaximumJsonUtf8Bytes} bytes.");
            }

            (_pipelineBytes ??= new List<byte>()).Add(item);
            return true;
        }

        if (_pipelineBytes != null)
        {
            throw MixedPipelineInput();
        }

        _nonBytePipelineInputSeen = true;
        return false;
    }

    /// <summary>Returns the reassembled JSON payload, or <see langword="null"/> when no scalar bytes were piped.</summary>
    protected byte[]? CompletePipelineBytes() => _pipelineBytes?.ToArray();

    /// <summary>Resolves typed, JSON, file, or ImagePlayground portable input into native editable Visio.</summary>
    protected OfficeVisioVisualConversionResult ResolveVisioVisual(object inputObject)
    {
        if (inputObject == null)
        {
            throw new PSArgumentNullException(nameof(inputObject));
        }

        PSObject input = PSObject.AsPSObject(inputObject);
        object value = input.BaseObject;
        if (value is OfficeVisioVisualConversionResult converted)
        {
            RejectConversionOverrides();
            return converted;
        }

        OfficeVisioVisualOptions options = CreateOptions();
        if (value is VisualArtifact artifact)
        {
            return artifact.ToOfficeVisio(options);
        }
        if (value is VisualArtifactInterchangeEnvelope envelope)
        {
            return envelope.ToOfficeVisio(options);
        }
        if (value is byte[] jsonBytes)
        {
            return jsonBytes.ToOfficeVisio(options);
        }
        if (value is FileInfo fileInfo)
        {
            return ConvertJsonFile(fileInfo.FullName, options);
        }
        if (value is PathInfo pathInfo)
        {
            return ConvertJsonFile(pathInfo.ProviderPath, options);
        }
        if (value is string text)
        {
            string trimmed = text.TrimStart('\uFEFF', ' ', '\t', '\r', '\n');
            if (trimmed.StartsWith("{", StringComparison.Ordinal))
            {
                return VisualArtifactInterchangeEnvelope.FromJson(trimmed).ToOfficeVisio(options);
            }
            return ConvertJsonFile(SessionState.Path.GetUnresolvedProviderPathFromPSPath(text), options);
        }
        if (input.TypeNames.Contains("ImagePlayground.VisualArtifact"))
        {
            object? payload = input.Properties["OfficeVisualInterchangeJson"]?.Value;
            if (payload is not byte[] portableBytes || portableBytes.Length == 0)
            {
                throw new PSArgumentException(
                    "ImagePlayground.VisualArtifact input must provide non-empty OfficeVisualInterchangeJson bytes. " +
                    "Recreate the artifact with a semantic-interchange capable ImagePlayground version.",
                    nameof(inputObject));
            }
            return portableBytes.ToOfficeVisio(options);
        }

        throw new PSArgumentException(
            "InputObject must be a ChartForgeX VisualArtifact, CFX interchange envelope or JSON bytes/file, " +
            "ImagePlayground.VisualArtifact, or OfficeVisioVisualConversionResult. Received " + value.GetType().FullName + ".",
            nameof(inputObject));
    }

    private OfficeVisioVisualConversionResult ConvertJsonFile(string path, OfficeVisioVisualOptions options)
    {
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"CFX interchange JSON file '{fullPath}' was not found.", fullPath);
        }
        if (!string.Equals(Path.GetExtension(fullPath), ".json", StringComparison.OrdinalIgnoreCase))
        {
            throw new PSArgumentException("Native Visio visual file input must be CFX interchange JSON with a .json extension.", nameof(path));
        }
        long length = new FileInfo(fullPath).Length;
        if (length > VisualArtifactInterchangeEnvelope.MaximumJsonUtf8Bytes)
        {
            throw new PSArgumentOutOfRangeException(
                nameof(path),
                length,
                $"CFX interchange UTF-8 JSON must not exceed {VisualArtifactInterchangeEnvelope.MaximumJsonUtf8Bytes} bytes.");
        }
        return File.ReadAllBytes(fullPath).ToOfficeVisio(options);
    }

    private OfficeVisioVisualOptions CreateOptions() => new OfficeVisioVisualOptions
    {
        PageName = PageName,
        PixelsPerInch = PixelsPerInch,
        UseNaturalPageSize = UseNaturalPageSize.IsPresent,
        IncludeTitle = !NoTitle.IsPresent,
        IncludeGroups = !NoGroups.IsPresent,
        IncludeShapeData = !NoShapeData.IsPresent,
        IncludeHyperlinks = !NoHyperlinks.IsPresent
    };

    private void RejectConversionOverrides()
    {
        foreach (string name in new[] { nameof(PageName), nameof(UseNaturalPageSize), nameof(PixelsPerInch), nameof(NoTitle), nameof(NoGroups), nameof(NoShapeData), nameof(NoHyperlinks) })
        {
            if (MyInvocation.BoundParameters.ContainsKey(name))
            {
                throw new PSArgumentException($"-{name} cannot be used with an existing OfficeVisioVisualConversionResult.", name);
            }
        }
    }

    private static PSArgumentException MixedPipelineInput() => new PSArgumentException(
        "A byte-stream pipeline cannot be mixed with typed artifacts, files, or JSON text in the same invocation.",
        "InputObject");
}
