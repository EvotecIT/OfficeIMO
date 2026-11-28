using System;
using System.Collections.Generic;

namespace OfficeIMO.Entities;

/// <summary>VAT schemes for the EU member states and Northern Ireland.</summary>
public static class EuVatSchemes {
    /// <summary>Registry in VIES prefix order. Greece uses EL and Northern Ireland uses XI.</summary>
    public static IReadOnlyList<EntityScheme> All => Schemes.All;

    /// <summary>Returns a scheme by two-character VIES prefix.</summary>
    public static EntityScheme Get(string prefix) {
        if (prefix == null) throw new ArgumentNullException(nameof(prefix));
        if (Schemes.ByPrefix.TryGetValue(prefix, out PrefixedVatScheme? scheme)) return scheme;
        throw new KeyNotFoundException("The VIES prefix is not registered.");
    }

    internal static bool TryGet(string prefix, out PrefixedVatScheme? scheme) => Schemes.ByPrefix.TryGetValue(prefix, out scheme);

    private static class Schemes {
        private static readonly PrefixedVatScheme[] Registry = {
            new AtVatScheme(), new BeVatScheme(), new BgVatScheme(), new CyVatScheme(), new CzVatScheme(),
            new DeVatScheme(), new DkVatScheme(), new EeVatScheme(), new ElVatScheme(), new EsVatScheme(),
            new FiVatScheme(), new FrVatScheme(), new HrVatScheme(), new HuVatScheme(), new IeVatScheme(),
            new ItVatScheme(), new LtVatScheme(), new LuVatScheme(), new LvVatScheme(), new MtVatScheme(),
            new NlVatScheme(), new PlVatScheme(), new PtVatScheme(), new RoVatScheme(), new SeVatScheme(),
            new SiVatScheme(), new SkVatScheme(), new UkVatScheme("XI", "EU.VAT.XI"),
        };

        internal static readonly IReadOnlyList<EntityScheme> All = Array.AsReadOnly<EntityScheme>(Registry);
        internal static readonly Dictionary<string, PrefixedVatScheme> ByPrefix = CreateLookup();

        private static Dictionary<string, PrefixedVatScheme> CreateLookup() {
            var lookup = new Dictionary<string, PrefixedVatScheme>(StringComparer.OrdinalIgnoreCase);
            foreach (PrefixedVatScheme scheme in Registry) lookup.Add(scheme.Prefix, scheme);
            return lookup;
        }
    }
}
