using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>Describes content that the native binary PowerPoint writer cannot encode.</summary>
    public sealed class LegacyPptWriteFinding {
        internal LegacyPptWriteFinding(LegacyPptFeature feature, string code, string description,
            int? slideIndex = null, int? shapeIndex = null) {
            Feature = feature;
            Code = code;
            Description = description;
            SlideIndex = slideIndex;
            ShapeIndex = shapeIndex;
        }

        /// <summary>Gets the capability-contract feature associated with the loss finding.</summary>
        public LegacyPptFeature Feature { get; }

        /// <summary>Gets a stable finding code.</summary>
        public string Code { get; }

        /// <summary>Gets the finding description.</summary>
        public string Description { get; }

        /// <summary>Gets the zero-based slide index, when applicable.</summary>
        public int? SlideIndex { get; }

        /// <summary>Gets the zero-based shape index, when applicable.</summary>
        public int? ShapeIndex { get; }

        /// <inheritdoc />
        public override string ToString() => SlideIndex.HasValue
            ? $"{Code} [slide {SlideIndex.Value + 1}{(ShapeIndex.HasValue ? $", shape {ShapeIndex.Value + 1}" : string.Empty)}]: {Description}"
            : $"{Code}: {Description}";
    }

    /// <summary>Reports whether a presentation fits the native binary writer's supported subset.</summary>
    public sealed class LegacyPptWritePreflightReport {
        internal LegacyPptWritePreflightReport(IReadOnlyList<LegacyPptWriteFinding> findings) {
            Findings = findings;
        }

        /// <summary>Gets known conversion-loss findings.</summary>
        public IReadOnlyList<LegacyPptWriteFinding> Findings { get; }

        /// <summary>Gets whether writing would omit known content.</summary>
        public bool HasConversionLoss => Findings.Count > 0;

        /// <summary>Gets whether the default loss policy permits writing.</summary>
        public bool CanWrite => !HasConversionLoss;
    }
}
