using System;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class GetOfficePdfRedactionPlanCommand {
    /// <summary>Input PDF paths with wildcard expansion.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Pattern", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Area", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true, Position = 0)]
    [SupportsWildcards]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Text setting.</summary>
    [Parameter(ParameterSetName = "Text", Mandatory = true)]
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true)]
    public string[] Text { get; set; } = Array.Empty<string>();
    /// <summary>CaseSensitive setting.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    public SwitchParameter CaseSensitive { get; set; }
    /// <summary>Password setting.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>IgnorePermissionRestrictions setting.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Native parsing, authentication and resource policy.</summary>
    [Parameter(ParameterSetName = "Text")]
    [Parameter(ParameterSetName = "TextLiteralPath")]
    [Parameter(ParameterSetName = "Pattern")]
    [Parameter(ParameterSetName = "PatternLiteralPath")]
    [Parameter(ParameterSetName = "Area")]
    [Parameter(ParameterSetName = "AreaLiteralPath")]
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Literal input PDF paths; accepts FileInfo.FullName.</summary>
    [Parameter(ParameterSetName = "TextLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "PatternLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "AreaLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("FullName", "PSPath")]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Pattern setting.</summary>
    [Parameter(ParameterSetName = "Pattern", Mandatory = true)]
    [Parameter(ParameterSetName = "PatternLiteralPath", Mandatory = true)]
    public string[] Pattern { get; set; } = Array.Empty<string>();
    /// <summary>Area setting.</summary>
    [Parameter(ParameterSetName = "Area", Mandatory = true)]
    [Parameter(ParameterSetName = "AreaLiteralPath", Mandatory = true)]
    public PdfRedactionArea[] Area { get; set; } = Array.Empty<PdfRedactionArea>();
    /// <summary>Page setting.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    public int Page { get; set; }
    /// <summary>X setting.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength X { get; set; }
    /// <summary>Y setting.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Y { get; set; }
    /// <summary>Width setting.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; }
    /// <summary>Height setting.</summary>
    [Parameter(ParameterSetName = "Rectangle", Mandatory = true)]
    [Parameter(ParameterSetName = "RectangleLiteralPath", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; }
    /// <summary>Label setting.</summary>
    [Parameter(ParameterSetName = "Rectangle")]
    [Parameter(ParameterSetName = "RectangleLiteralPath")]
    public string? Label { get; set; }
}
