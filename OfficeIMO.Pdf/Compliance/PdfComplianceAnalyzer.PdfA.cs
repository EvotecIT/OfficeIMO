namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {

    private static void AddPdfARequirements(List<PdfComplianceRequirement> requirements, PdfComplianceProfile profile, PdfOptions options, PdfStandardFont[]? generatedStandardFonts, PdfGeneratedFontComplianceEvidence[]? generatedFontUsages) {
        PdfAIdentification? identification = options.PdfAIdentification;
        (int Part, string? Conformance) target = GetPdfAIdentificationTarget(profile);

        Add(requirements, "xmp-metadata", "Catalog XMP metadata",
            options.IncludeXmpMetadata || identification != null,
            "Catalog XMP metadata will be emitted.",
            "Enable PdfOptions.IncludeXmpMetadata or set PdfOptions.PdfAIdentification.");

        bool hasMatchingIdentification = identification != null &&
            identification.Part == target.Part &&
            IsPdfAConformanceMatch(identification.Conformance, target.Conformance);
        string expectedIdentification = target.Conformance == null
            ? "PDF/A-" + target.Part.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : "PDF/A-" + target.Part.ToString(System.Globalization.CultureInfo.InvariantCulture) + target.Conformance!.ToLowerInvariant();
        Add(requirements, "pdfa-identification", "PDF/A identification XMP",
            hasMatchingIdentification,
            "PDF/A identification metadata matches " + expectedIdentification + ".",
            "Set PdfOptions.SetPdfAIdentification(...) to " + expectedIdentification + " before claiming this profile.");

        requirements.Add(BuildPdfAEncryptionPolicyRequirement(options));

        Add(requirements, "output-intent", "Catalog output intent",
            options.OutputIntent != null,
            "A catalog output intent with a parseable RGB, GRAY, or CMYK ICC profile is configured.",
            "Set PdfOptions.OutputIntent or PdfDocument.OutputIntent(...) with ICC profile bytes that pass OfficeIMO's ICC header checks.");

        requirements.Add(BuildOutputIntentPolicyRequirement(options));

        AddEmbeddedFontCoverageRequirement(requirements, options, generatedStandardFonts, generatedFontUsages);

        requirements.Add(new PdfComplianceRequirement(
            "verapdf-validation",
            "veraPDF validation evidence",
            PdfComplianceRequirementStatus.Unsupported,
            "The optional veraPDF test gate exists for groundwork fixtures, but profile success has not been enabled for generated output."));
    }

    private static PdfComplianceRequirement BuildPdfAEncryptionPolicyRequirement(PdfOptions options) {
        if (options.EncryptionSnapshot == null) {
            return new PdfComplianceRequirement(
                "pdfa-no-encryption",
                "PDF/A encryption policy",
                PdfComplianceRequirementStatus.Satisfied,
                "No Standard security encryption is configured for the PDF/A-backed output.");
        }

        return new PdfComplianceRequirement(
            "pdfa-no-encryption",
            "PDF/A encryption policy",
            PdfComplianceRequirementStatus.Missing,
            "PDF/A-backed profiles cannot be claimed with Standard security encryption. Clear PdfOptions.Encryption before assessing or claiming PDF/A or e-invoice readiness.");
    }

    private static PdfComplianceRequirement BuildOutputIntentPolicyRequirement(PdfOptions options) {
        PdfOutputIntent? outputIntent = options.OutputIntent;
        if (outputIntent == null) {
            return new PdfComplianceRequirement(
                "output-intent-policy",
                "Profile-specific output-intent policy",
                PdfComplianceRequirementStatus.Missing,
                "Configure a catalog output intent before checking profile-specific output-intent policy.");
        }

        if (outputIntent.Policy == PdfOutputIntentPolicy.Unspecified) {
            return new PdfComplianceRequirement(
                "output-intent-policy",
                "Profile-specific output-intent policy",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOutputIntent.Policy or pass a PdfOutputIntentPolicy value to SetOutputIntent/OutputIntent so PDF/A readiness can distinguish generic ICC output from a known profile policy.");
        }

        if (outputIntent.Policy == PdfOutputIntentPolicy.SrgbIec6196621) {
            if (outputIntent.ColorComponents != 3) {
                return new PdfComplianceRequirement(
                    "output-intent-policy",
                    "Profile-specific output-intent policy",
                    PdfComplianceRequirementStatus.Missing,
                    "The sRGB IEC61966-2.1 output-intent policy requires an RGB ICC profile.");
            }

            if (!string.Equals(outputIntent.OutputConditionIdentifier, PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, StringComparison.Ordinal)) {
                return new PdfComplianceRequirement(
                    "output-intent-policy",
                    "Profile-specific output-intent policy",
                    PdfComplianceRequirementStatus.Missing,
                    "The sRGB IEC61966-2.1 output-intent policy requires OutputConditionIdentifier to be " + PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier + ".");
            }

            return new PdfComplianceRequirement(
                "output-intent-policy",
                "Profile-specific output-intent policy",
                PdfComplianceRequirementStatus.Satisfied,
                "The output intent declares the sRGB IEC61966-2.1 policy with RGB ICC profile evidence. External veraPDF validation is still required before claiming PDF/A conformance.");
        }

        return new PdfComplianceRequirement(
            "output-intent-policy",
            "Profile-specific output-intent policy",
            PdfComplianceRequirementStatus.Missing,
            "Use a supported PDF output-intent policy.");
    }

    private static void AddEmbeddedFontCoverageRequirement(List<PdfComplianceRequirement> requirements, PdfOptions options, PdfStandardFont[]? generatedStandardFonts, PdfGeneratedFontComplianceEvidence[]? generatedFontUsages) {
        if (generatedStandardFonts == null) {
            requirements.Add(new PdfComplianceRequirement(
                "embedded-font-coverage",
                "Embedded font coverage",
                PdfComplianceRequirementStatus.Unsupported,
                "OfficeIMO.Pdf can embed caller-supplied TrueType and OpenType/CFF files for standard-font slots, but generated-document font usage was not supplied for this readiness assessment."));
            return;
        }

        if (generatedStandardFonts.Length == 0) {
            requirements.Add(new PdfComplianceRequirement(
                "embedded-font-coverage",
                "Embedded font coverage",
                PdfComplianceRequirementStatus.Satisfied,
                "No generated standard-font resources were reported for this document."));
            return;
        }

        var missingFonts = new List<PdfStandardFont>();
        var invalidFonts = new List<string>();
        if (generatedFontUsages != null && generatedFontUsages.Length > 0) {
            for (int i = 0; i < generatedFontUsages.Length; i++) {
                PdfGeneratedFontComplianceEvidence usage = generatedFontUsages[i];
                IReadOnlyDictionary<PdfStandardFont, PdfEmbeddedFont> scopedEmbeddedFonts = usage.Options.EmbeddedFonts;
                if (!scopedEmbeddedFonts.TryGetValue(usage.Font, out PdfEmbeddedFont? embeddedFont)) {
                    AddMissingFont(missingFonts, usage.Font);
                    continue;
                }

                if (!TryParseEmbeddedFont(embeddedFont, out string? invalidReason)) {
                    AddInvalidFont(invalidFonts, usage.Font, invalidReason);
                }
            }
        } else {
            IReadOnlyDictionary<PdfStandardFont, PdfEmbeddedFont> embeddedFonts = options.EmbeddedFonts;
            for (int i = 0; i < generatedStandardFonts.Length; i++) {
                PdfStandardFont font = generatedStandardFonts[i];
                if (!embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? embeddedFont)) {
                    AddMissingFont(missingFonts, font);
                    continue;
                }

                if (!TryParseEmbeddedFont(embeddedFont, out string? invalidReason)) {
                    AddInvalidFont(invalidFonts, font, invalidReason);
                }
            }
        }

        if (missingFonts.Count == 0 && invalidFonts.Count == 0) {
            requirements.Add(new PdfComplianceRequirement(
                "embedded-font-coverage",
                "Embedded font coverage",
                PdfComplianceRequirementStatus.Satisfied,
                "Every generated standard-font resource has a parseable embedded TrueType or OpenType/CFF mapping."));
            return;
        }

        var diagnostics = new List<string>();
        if (missingFonts.Count > 0) {
            diagnostics.Add("embed TrueType or OpenType/CFF mappings for generated standard-font resources: " + string.Join(", ", missingFonts.Select(font => font.ToBaseFontName()).ToArray()));
        }

        if (invalidFonts.Count > 0) {
            diagnostics.Add("replace invalid embedded TrueType or OpenType/CFF mappings: " + string.Join(", ", invalidFonts.ToArray()));
        }

        requirements.Add(new PdfComplianceRequirement(
            "embedded-font-coverage",
            "Embedded font coverage",
            PdfComplianceRequirementStatus.Missing,
            char.ToUpperInvariant(diagnostics[0][0]) + diagnostics[0].Substring(1) + (diagnostics.Count > 1 ? "; " + string.Join("; ", diagnostics.Skip(1).ToArray()) : string.Empty) + "."));
    }

    private static void AddMissingFont(List<PdfStandardFont> missingFonts, PdfStandardFont font) {
        if (!missingFonts.Contains(font)) {
            missingFonts.Add(font);
        }
    }

    private static void AddInvalidFont(List<string> invalidFonts, PdfStandardFont font, string? invalidReason) {
        string diagnostic = font.ToBaseFontName() + " (" + invalidReason + ")";
        if (!invalidFonts.Contains(diagnostic)) {
            invalidFonts.Add(diagnostic);
        }
    }

    private static bool TryParseEmbeddedFont(PdfEmbeddedFont embeddedFont, out string? invalidReason) {
        try {
            if (IsOpenTypeCffFontData(embeddedFont.Data)) {
                PdfOpenTypeCffFontProgram.Parse(embeddedFont.Data, embeddedFont.FontName);
            } else {
                PdfTrueTypeFontProgram.Parse(embeddedFont.Data, embeddedFont.FontName);
            }

            invalidReason = null;
            return true;
        } catch (Exception ex) when (PdfFontDiagnostics.IsFontProgramException(ex)) {
            invalidReason = ex.Message;
            return false;
        }
    }

    private static bool IsOpenTypeCffFontData(byte[] fontData) =>
        fontData.Length >= 4 &&
        fontData[0] == 0x4F &&
        fontData[1] == 0x54 &&
        fontData[2] == 0x54 &&
        fontData[3] == 0x4F;

    private static (int Part, string? Conformance) GetPdfAIdentificationTarget(PdfComplianceProfile profile) {
        switch (profile) {
            case PdfComplianceProfile.PdfA2B:
                return (2, "B");
            case PdfComplianceProfile.PdfA2U:
                return (2, "U");
            case PdfComplianceProfile.PdfA2A:
                return (2, "A");
            case PdfComplianceProfile.PdfA3B:
                return (3, "B");
            case PdfComplianceProfile.PdfA3U:
                return (3, "U");
            case PdfComplianceProfile.PdfA3A:
                return (3, "A");
            case PdfComplianceProfile.PdfA4:
                return (4, string.Empty);
            case PdfComplianceProfile.PdfA4E:
                return (4, "E");
            case PdfComplianceProfile.PdfA4F:
                return (4, "F");
            case PdfComplianceProfile.FacturX:
            case PdfComplianceProfile.Zugferd:
                return (3, null);
            default:
                return (0, null);
        }
    }

    private static bool IsPdfAConformanceMatch(string? actual, string? expected) {
        if (expected == null) {
            return true;
        }

        if (expected.Length == 0) {
            return string.IsNullOrEmpty(actual);
        }

        return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
    }
}
