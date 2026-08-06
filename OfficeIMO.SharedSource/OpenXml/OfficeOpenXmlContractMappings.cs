using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.OpenXml.Internal;

internal static class OfficeOpenXmlContractMappings {
    internal static OpenSettings ToOpenXml(this OfficeOpenXmlLoadSettings? settings) {
        OfficeOpenXmlLoadSettings resolved = settings ?? new OfficeOpenXmlLoadSettings();
        return new OpenSettings {
            AutoSave = false,
            CompatibilityLevel = resolved.CompatibilityLevel switch {
                OfficeOpenXmlCompatibilityLevel.Default => CompatibilityLevel.Default,
                OfficeOpenXmlCompatibilityLevel.Version220 => CompatibilityLevel.Version_2_20,
                OfficeOpenXmlCompatibilityLevel.Version30 => CompatibilityLevel.Version_3_0,
                _ => throw new ArgumentOutOfRangeException(nameof(settings), resolved.CompatibilityLevel, "Unsupported Open XML compatibility level.")
            },
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                resolved.MarkupCompatibilityMode switch {
                    OfficeOpenXmlMarkupCompatibilityMode.NoProcess => MarkupCompatibilityProcessMode.NoProcess,
                    OfficeOpenXmlMarkupCompatibilityMode.ProcessLoadedPartsOnly => MarkupCompatibilityProcessMode.ProcessLoadedPartsOnly,
                    OfficeOpenXmlMarkupCompatibilityMode.ProcessAllParts => MarkupCompatibilityProcessMode.ProcessAllParts,
                    _ => throw new ArgumentOutOfRangeException(nameof(settings), resolved.MarkupCompatibilityMode, "Unsupported markup-compatibility mode.")
                },
                resolved.MarkupCompatibilityTargetVersion.ToOpenXml()),
            MaxCharactersInPart = resolved.MaxCharactersInPart
        };
    }

    internal static FileFormatVersions ToOpenXml(this OfficeOpenXmlFileFormatVersion version) =>
        version switch {
            OfficeOpenXmlFileFormatVersion.Office2007 => FileFormatVersions.Office2007,
            OfficeOpenXmlFileFormatVersion.Office2010 => FileFormatVersions.Office2010,
            OfficeOpenXmlFileFormatVersion.Office2013 => FileFormatVersions.Office2013,
            OfficeOpenXmlFileFormatVersion.Office2016 => FileFormatVersions.Office2016,
            OfficeOpenXmlFileFormatVersion.Office2019 => FileFormatVersions.Office2019,
            OfficeOpenXmlFileFormatVersion.Office2021 => FileFormatVersions.Office2021,
            OfficeOpenXmlFileFormatVersion.Microsoft365 => FileFormatVersions.Microsoft365,
            _ => throw new ArgumentOutOfRangeException(nameof(version), version, "Unsupported Open XML file-format version.")
        };

    internal static OfficeOpenXmlValidationError ToOfficeValidationError(this ValidationErrorInfo error) {
        if (error == null) throw new ArgumentNullException(nameof(error));
        return new OfficeOpenXmlValidationError(
            error.Id,
            error.ErrorType switch {
                ValidationErrorType.Schema => OfficeOpenXmlValidationErrorType.Schema,
                ValidationErrorType.Semantic => OfficeOpenXmlValidationErrorType.Semantic,
                ValidationErrorType.Package => OfficeOpenXmlValidationErrorType.Package,
                ValidationErrorType.MarkupCompatibility => OfficeOpenXmlValidationErrorType.MarkupCompatibility,
                _ => throw new ArgumentOutOfRangeException(nameof(error), error.ErrorType, "Unsupported Open XML validation error type.")
            },
            error.Description,
            error.Path?.XPath,
            error.Part?.Uri.ToString(),
            error.Node?.LocalName,
            error.RelatedPart?.Uri.ToString(),
            error.RelatedNode?.LocalName);
    }
}
