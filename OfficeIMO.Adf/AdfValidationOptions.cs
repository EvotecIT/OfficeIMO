namespace OfficeIMO.Adf;

/// <summary>Selects the validation contract independently of source preservation.</summary>
public enum AdfValidationProfile {
    /// <summary>Checks recognized relationships while retaining newer node and mark types with warnings.</summary>
    ForwardCompatible,
    /// <summary>Checks the complete bundled Atlassian ADF JSON schema, including required properties and attributes.</summary>
    FullSchema
}

/// <summary>Options for structural or pinned-schema validation.</summary>
public sealed class AdfValidationOptions : AdfProcessingOptions {
    /// <summary>Validation profile. Defaults to forward-compatible structural validation.</summary>
    public AdfValidationProfile Profile { get; set; }

    /// <summary>Optional caller-defined destination product/version restrictions, applied after the selected validation contract.</summary>
    /// <remarks>Validation reports up to 1,000 destination errors without modifying source nodes or marks.</remarks>
    public AdfDestinationPolicy? DestinationPolicy { get; set; }

    /// <summary>Maximum rule evaluations for full-schema validation. Defaults to 100 million.</summary>
    /// <remarks>Exceeding this positive bound produces an ADF_SCHEMA_LIMIT issue. Structural validation uses the document graph limits.</remarks>
    public long MaximumSchemaEvaluations { get; set; } = 100000000;
}
