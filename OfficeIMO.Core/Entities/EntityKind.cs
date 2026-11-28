using System;

namespace OfficeIMO.Entities;

/// <summary>A case-insensitive entity category with an open set of names.</summary>
public readonly struct EntityKind : IEquatable<EntityKind> {
    private readonly string? _name;

    /// <summary>Creates a category for use by a scheme and its consumers.</summary>
    public EntityKind(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("An entity kind is required.", nameof(name));
        _name = name;
    }

    /// <summary>Email address.</summary>
    public static EntityKind Email { get; } = new("Email");
    /// <summary>Uniform resource locator.</summary>
    public static EntityKind Url { get; } = new("Url");
    /// <summary>Telephone number.</summary>
    public static EntityKind Phone { get; } = new("Phone");
    /// <summary>Postal address.</summary>
    public static EntityKind PostalAddress { get; } = new("PostalAddress");
    /// <summary>Company name.</summary>
    public static EntityKind CompanyName { get; } = new("CompanyName");
    /// <summary>Bank account identifier.</summary>
    public static EntityKind BankAccount { get; } = new("BankAccount");
    /// <summary>Bank or branch identifier.</summary>
    public static EntityKind BankIdentifier { get; } = new("BankIdentifier");
    /// <summary>Payment card number.</summary>
    public static EntityKind PaymentCard { get; } = new("PaymentCard");
    /// <summary>Financial security identifier.</summary>
    public static EntityKind SecurityId { get; } = new("SecurityId");
    /// <summary>Legal entity identifier.</summary>
    public static EntityKind LegalEntityId { get; } = new("LegalEntityId");
    /// <summary>National personal identifier.</summary>
    public static EntityKind NationalId { get; } = new("NationalId");
    /// <summary>Taxpayer identifier.</summary>
    public static EntityKind TaxId { get; } = new("TaxId");
    /// <summary>Value-added tax registration identifier.</summary>
    public static EntityKind VatId { get; } = new("VatId");
    /// <summary>Company registration identifier.</summary>
    public static EntityKind CompanyId { get; } = new("CompanyId");
    /// <summary>Identity document number.</summary>
    public static EntityKind IdentityDocument { get; } = new("IdentityDocument");

    /// <inheritdoc />
    public bool Equals(EntityKind other) => StringComparer.OrdinalIgnoreCase.Equals(_name, other._name);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is EntityKind other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => _name == null ? 0 : StringComparer.OrdinalIgnoreCase.GetHashCode(_name);
    /// <inheritdoc />
    public override string ToString() => _name ?? string.Empty;
    /// <summary>Compares category names without regard to case.</summary>
    public static bool operator ==(EntityKind left, EntityKind right) => left.Equals(right);
    /// <summary>Compares category names without regard to case.</summary>
    public static bool operator !=(EntityKind left, EntityKind right) => !left.Equals(right);
}
