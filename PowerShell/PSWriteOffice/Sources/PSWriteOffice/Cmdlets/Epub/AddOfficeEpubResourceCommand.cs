using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Epub;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Epub;

/// <summary>Adds an image, font, stylesheet or media payload without executing it.</summary>
/// <example>
///   <summary>Adds an image, font, stylesheet or media payload without executing it.</summary>
///   <code>$book | Add-OfficeEpubResource -Path ./cover.png -Id cover -MediaType image/png</code>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeEpubResource")]
[OutputType(typeof(EpubManifestItem))]
public sealed class AddOfficeEpubResourceCommand : PSCmdlet {
    /// <summary>Live editable publication.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EpubPublication Publication { get; set; } = null!;

    /// <summary>Literal payload path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Unique manifest identifier.</summary>
    [Parameter(Mandatory = true)]
    public string Id { get; set; } = string.Empty;

    /// <summary>Explicit MIME type.</summary>
    [Parameter(Mandatory = true)]
    public string MediaType { get; set; } = string.Empty;

    /// <summary>Optional package-relative path; omission uses EPUB/assets and the source filename.</summary>
    [Parameter]
    public string? ContainerPath { get; set; }

    /// <summary>Maximum input payload bytes.</summary>
    [Parameter]
    [ValidateRange(1L, int.MaxValue)]
    public long MaxInputBytes { get; set; } = 33554432;

    /// <summary>Return the created or changed object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var path = OfficeFileCommandUtilities.ResolvePath(this, Path);
        using var source = File.OpenRead(path);
        if (source.Length > MaxInputBytes) throw new InvalidDataException("The resource exceeds MaxInputBytes.");
        byte[] payload = OfficeStreamReader.ReadAllBytes(source, MaxInputBytes);
        var resource = Publication.AddResource(Id, ContainerPath ?? "EPUB/assets/" + System.IO.Path.GetFileName(path), MediaType, payload);
        if (PassThru.IsPresent) WriteObject(resource);
    }
}
