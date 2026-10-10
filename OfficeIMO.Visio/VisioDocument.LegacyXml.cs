using System;
using System.IO;
using System.IO.Packaging;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    /// <summary>Imports a Visio 2002/2003 VDX drawing, VSX stencil, or VTX template into the editable model with a fidelity report.</summary>
    /// <remarks>This explicit conversion does not associate the XML source with ordinary Save, which writes Open XML packages.</remarks>
    public static OfficeConversionResult<VisioDocument, VisioXmlConversionReport> LoadLegacyXml(string path, VisioLoadOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        string extension = Path.GetExtension(path);
        VisioPackageType family = extension.Equals(".vsx", StringComparison.OrdinalIgnoreCase) ? VisioPackageType.Stencil
            : extension.Equals(".vtx", StringComparison.OrdinalIgnoreCase) ? VisioPackageType.Template : VisioPackageType.Drawing;
        using var stream = File.OpenRead(path);
        return LoadLegacyXml(stream, family, options);
    }

    /// <summary>Imports legacy XML from a caller-owned stream. Specify the family because legacy XML has no reliable drawing/template discriminator.</summary>
    public static OfficeConversionResult<VisioDocument, VisioXmlConversionReport> LoadLegacyXml(Stream stream,
        VisioPackageType packageType = VisioPackageType.Drawing, VisioLoadOptions? options = null, CancellationToken cancellationToken = default) {
        VisioLegacyXmlCodec.ValidateFamily(packageType);
        VisioLoadOptions resolved = options ?? new VisioLoadOptions();
        if (resolved.MaxLegacyXmlCharacters < 1 || resolved.MaxLegacyXmlDepth < 1 || resolved.MaxLegacyXmlElements < 1)
            throw new ArgumentOutOfRangeException(nameof(options), "Legacy XML limits must be positive.");
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, ResolveInputLimit(resolved));
        using var source = new MemoryStream(bytes, writable: false);
        using var reader = XmlReader.Create(source, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = resolved.MaxLegacyXmlCharacters
        });
        using var bounded = new OfficeXmlLimitingReader(reader, "Visio XML", resolved.MaxLegacyXmlDepth,
            resolved.MaxLegacyXmlElements, resolved.MaxLegacyXmlElements, cancellationToken);
        XDocument xml = XDocument.Load(bounded, LoadOptions.PreserveWhitespace);
        var report = new VisioXmlConversionReport();
        using MemoryStream normalized = VisioLegacyXmlCodec.ToPackage(xml, packageType, report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] packageBytes = normalized.ToArray();
        ValidatePackageSecurity(packageBytes, resolved);
        using Package package = Package.Open(normalized, FileMode.Open, FileAccess.Read);
        VisioDocument document = LoadCore(package, null, cancellationToken);
        report.Add("VDX_MODEL_PROFILE", "The existing Visio model interprets supported shapes, text, masters and connectors; formula evaluation and advanced rendering follow its documented profile.", OfficeConversionLossKind.Approximation);
        cancellationToken.ThrowIfCancellationRequested();
        return new OfficeConversionResult<VisioDocument, VisioXmlConversionReport>(document, report);
    }

    /// <summary>Encodes the current drawing, stencil or template as Visio 2003 XML with operation-level loss diagnostics.</summary>
    public OfficeConversionResult<byte[], VisioXmlConversionReport> ToLegacyXmlResult() {
        VisioLegacyXmlCodec.ValidateFamily(_packageType);
        if (HasVbaProject) throw new NotSupportedException("Legacy XML export does not carry VBA projects.");
        ThrowIfInvalidForSave();
        var report = new VisioXmlConversionReport();
        using MemoryStream bytes = CreatePackageStream();
        using Package package = Package.Open(bytes, FileMode.Open, FileAccess.Read);
        XDocument xml = VisioLegacyXmlCodec.FromPackage(package, report);
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false), Indent = false })) xml.Save(writer);
        return new OfficeConversionResult<byte[], VisioXmlConversionReport>(output.ToArray(), report);
    }

    /// <summary>Writes a legacy XML copy atomically. Omitted content requires explicit allowOmissions; inspect the returned report for approximations.</summary>
    public VisioXmlConversionReport SaveLegacyXml(string path, bool allowOmissions = false) {
        var result = ToLegacyXmlResult();
        RequireLegacyXmlOutput(result.Report, allowOmissions);
        OfficeFileCommit.WriteAllBytes(path, result.Value);
        return result.Report;
    }
    /// <summary>Writes legacy XML once to a caller-owned stream without changing the associated destination.</summary>
    public VisioXmlConversionReport SaveLegacyXml(Stream stream, bool allowOmissions = false) {
        var result = ToLegacyXmlResult();
        RequireLegacyXmlOutput(result.Report, allowOmissions);
        OfficeStreamWriter.WriteAllBytes(stream, result.Value);
        return result.Report;
    }
    private static void RequireLegacyXmlOutput(VisioXmlConversionReport report, bool allowOmissions) {
        if (!allowOmissions && report.HasOmissions) throw new OfficeConversionException("Legacy XML export would omit content. Inspect ToLegacyXmlResult before permitting omissions.", report);
    }
}
