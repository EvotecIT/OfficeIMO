namespace OfficeIMO.Invoicing;

/// <summary>Controls intentional information reduction required by lower Factur-X profiles.</summary>
public enum InvoiceProjectionPolicy {
    /// <summary>Reject output whenever the selected profile cannot carry populated semantic data.</summary>
    RejectDataLoss,
    /// <summary>Allow only the documented, profile-defined reduction and return it from <see cref="InvoiceSerializer.InspectTarget"/> as warnings.</summary>
    AllowProfileDefinedDataLoss
}

/// <summary>Explicit specification release, syntax, and profile for deterministic XML emission.</summary>
public sealed class InvoiceXmlOptions {
    /// <summary>Creates an exact output contract. All three dimensions are required and compatibility is checked immediately.</summary>
    public InvoiceXmlOptions(InvoiceSpecificationRelease release, InvoiceSyntax syntax, InvoiceProfile profile,
        InvoiceProjectionPolicy projectionPolicy = InvoiceProjectionPolicy.RejectDataLoss) {
        InvoiceSpecificationContracts.Validate(release, syntax, profile);
        if (projectionPolicy < InvoiceProjectionPolicy.RejectDataLoss || projectionPolicy > InvoiceProjectionPolicy.AllowProfileDefinedDataLoss)
            throw new ArgumentOutOfRangeException(nameof(projectionPolicy));
        Release = release;
        Syntax = syntax;
        Profile = profile;
        ProjectionPolicy = projectionPolicy;
    }
    /// <summary>Specification and rules release.</summary>
    public InvoiceSpecificationRelease Release { get; }
    /// <summary>Output syntax.</summary>
    public InvoiceSyntax Syntax { get; }
    /// <summary>Output guideline.</summary>
    public InvoiceProfile Profile { get; }
    /// <summary>Policy for intentional data reduction required by lower Factur-X profiles.</summary>
    public InvoiceProjectionPolicy ProjectionPolicy { get; }

    /// <summary>Returns the supported immutable authoring targets from the same contract owner used by the constructor.</summary>
    public static IReadOnlyList<InvoiceXmlOptions> GetSupportedTargets(InvoiceProjectionPolicy projectionPolicy = InvoiceProjectionPolicy.RejectDataLoss) {
        if (projectionPolicy < InvoiceProjectionPolicy.RejectDataLoss || projectionPolicy > InvoiceProjectionPolicy.AllowProfileDefinedDataLoss)
            throw new ArgumentOutOfRangeException(nameof(projectionPolicy));
        var targets = new List<InvoiceXmlOptions>();
        foreach (InvoiceSpecificationRelease release in Enum.GetValues(typeof(InvoiceSpecificationRelease)))
            foreach (InvoiceSyntax syntax in Enum.GetValues(typeof(InvoiceSyntax)))
                foreach (InvoiceProfile profile in Enum.GetValues(typeof(InvoiceProfile)))
                    if (InvoiceSpecificationContracts.IsSupported(release, syntax, profile)) targets.Add(new(release, syntax, profile, projectionPolicy));
        return targets.AsReadOnly();
    }
}
