using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>Legal Entity Identifier per ISO 17442-1 and the <see href="https://www.leiroc.org/lei.htm">LEI Regulatory Oversight Committee</see>.</summary>
internal sealed class LeiScheme : CompactEntityScheme {
    public override string Id => "LEI";
    public override EntityKind Kind => EntityKind.LegalEntityId;
    public override string DisplayName => "Legal Entity Identifier";
    public override string Issuer => "Global Legal Entity Identifier Foundation";
    public override Uri Specification { get; } = new("https://www.leiroc.org/lei.htm");
    public override int MinLength => 20;
    public override int MaxLength => 20;
    public override EntityCheckKind Check => EntityCheckKind.Mod97_10;
    public override EntitySensitivity Sensitivity => EntitySensitivity.Organizational;

    protected override EntityValidationResult ValidateCompact(string value) {
        if (value[4] != '0' || value[5] != '0' || !EntityCharacters.IsDigits(value, 18, 2)) return Result(EntityValidationStatus.InvalidFormat);
        return Checked(value, CheckDigits.Mod97(value) == 1);
    }
}
