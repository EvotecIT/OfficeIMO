using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.PowerPoint {
    /// <summary>Chooses the reports produced by the unified presentation inspection workflow.</summary>
    public sealed class PowerPointInspectionOptions {
        /// <summary>Runs Open XML package validation.</summary>
        public bool ValidatePackage { get; set; } = true;

        /// <summary>Runs layout, text-fit, collision, asset, and visual-snapshot preflight.</summary>
        public bool InspectPreflight { get; set; } = true;

        /// <summary>Runs the accessibility policy inspection.</summary>
        public bool InspectAccessibility { get; set; } = true;

        /// <summary>Inventories editable, preserved, partially editable, and unsupported features.</summary>
        public bool InspectFeatures { get; set; }

        /// <summary>Inspects classic and modern review comments.</summary>
        public bool InspectReviewComments { get; set; }

        /// <summary>Inspects animation and timing markup.</summary>
        public bool InspectAnimations { get; set; }

        /// <summary>Inspects package signature metadata.</summary>
        public bool InspectSignatures { get; set; }

        /// <summary>Creates structural, extraction, accessibility, PNG, SVG, and snapshot proof.</summary>
        public bool InspectVisuals { get; set; }

        /// <summary>Open XML version used by package validation.</summary>
        public FileFormatVersions FileFormatVersion { get; set; } = FileFormatVersions.Microsoft365;

        /// <summary>Optional preflight policy.</summary>
        public PowerPointDeckPreflightOptions? Preflight { get; set; }

        /// <summary>Optional accessibility policy.</summary>
        public PowerPointAccessibilityOptions? Accessibility { get; set; }

        /// <summary>Source label recorded by visual proof.</summary>
        public string VisualSourceKind { get; set; } = "generated";
    }

    /// <summary>Unified inspection result over the same presentation model used for authoring and editing.</summary>
    public sealed class PowerPointInspectionReport {
        internal PowerPointInspectionReport(IList<ValidationErrorInfo> packageErrors,
            PowerPointDeckPreflightReport? preflight, PowerPointAccessibilityReport? accessibility,
            PowerPointFeatureReport? features, PowerPointReviewReport? reviewComments,
            PowerPointAnimationReport? animations, PowerPointSignatureReport? signatures,
            PowerPointVisualProofReport? visuals) {
            PackageErrors = new ReadOnlyCollection<ValidationErrorInfo>(
                new List<ValidationErrorInfo>(packageErrors ?? throw new ArgumentNullException(nameof(packageErrors))));
            Preflight = preflight;
            Accessibility = accessibility;
            Features = features;
            ReviewComments = reviewComments;
            Animations = animations;
            Signatures = signatures;
            Visuals = visuals;
        }

        /// <summary>Open XML package validation errors.</summary>
        public IReadOnlyList<ValidationErrorInfo> PackageErrors { get; }

        /// <summary>Layout and visual preflight, when requested.</summary>
        public PowerPointDeckPreflightReport? Preflight { get; }

        /// <summary>Accessibility inspection, when requested.</summary>
        public PowerPointAccessibilityReport? Accessibility { get; }

        /// <summary>Feature inventory, when requested.</summary>
        public PowerPointFeatureReport? Features { get; }

        /// <summary>Review-comment inspection, when requested.</summary>
        public PowerPointReviewReport? ReviewComments { get; }

        /// <summary>Animation inspection, when requested.</summary>
        public PowerPointAnimationReport? Animations { get; }

        /// <summary>Signature inspection, when requested.</summary>
        public PowerPointSignatureReport? Signatures { get; }

        /// <summary>Rendered visual proof, when requested.</summary>
        public PowerPointVisualProofReport? Visuals { get; }

        /// <summary>Whether requested package, preflight, accessibility, and visual checks found no errors.</summary>
        public bool IsSuccessful => PackageErrors.Count == 0 &&
            (Preflight == null || Preflight.IsSuccessful) &&
            (Accessibility == null || Accessibility.IsSuccessful) &&
            (Visuals == null || Visuals.IsSuccessful);
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Runs a coherent set of inspections over the current presentation without changing it.
        /// </summary>
        public PowerPointInspectionReport Inspect(PowerPointInspectionOptions? options = null) {
            ThrowIfDisposed();
            PowerPointInspectionOptions resolved = options ?? new PowerPointInspectionOptions();
            List<ValidationErrorInfo> packageErrors = resolved.ValidatePackage
                ? ValidateDocument(resolved.FileFormatVersion)
                : new List<ValidationErrorInfo>();
            PowerPointDeckPreflightReport? preflight = resolved.InspectPreflight
                ? InspectPreflight(resolved.Preflight)
                : null;
            PowerPointAccessibilityReport? accessibility = resolved.InspectAccessibility
                ? InspectAccessibility(resolved.Accessibility)
                : null;
            PowerPointFeatureReport? features = resolved.InspectFeatures ? InspectFeatures() : null;
            PowerPointReviewReport? reviewComments = resolved.InspectReviewComments ? InspectReviewComments() : null;
            PowerPointAnimationReport? animations = resolved.InspectAnimations ? InspectAnimations() : null;
            PowerPointSignatureReport? signatures = resolved.InspectSignatures ? InspectSignatures() : null;
            PowerPointVisualProofReport? visuals = resolved.InspectVisuals
                ? InspectVisuals(resolved.VisualSourceKind)
                : null;
            return new PowerPointInspectionReport(packageErrors, preflight, accessibility, features, reviewComments,
                animations, signatures, visuals);
        }
    }
}
