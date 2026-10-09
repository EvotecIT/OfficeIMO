namespace OfficeIMO.Workflows;

/// <summary>Caller-supplied organization responsible for compliance testing and certification. OfficeIMO does not verify the claim.</summary>
/// <param name="Name">Certifier name (90), at most 4096 characters.</param>
/// <param name="Url">Absolute HTTP(S) certifier or certification-scheme page (93), without credentials.</param>
public sealed record BookOnixAccessibilityCertification(string Name, string Url) {
    /// <summary>Optional name of the organization credentialling the certifier (88), at most 4096 characters.</summary>
    public string? CredentiallingOrganizationName { get; init; }
    /// <summary>Optional absolute HTTP(S) page belonging to the credentialling organization (89), without credentials.</summary>
    public string? CredentiallingOrganizationUrl { get; init; }
}
