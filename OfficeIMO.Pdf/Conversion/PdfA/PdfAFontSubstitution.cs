namespace OfficeIMO.Pdf;

/// <summary>Chooses where PDF/A conversion may find a font program for a font the source does not embed.</summary>
/// <remarks>
/// A substitute is embedded only when every glyph the pages draw exists in it with the advance width the page already
/// positions it by, and only when its OS/2 embedding permission allows embedding. Supplied and installed faces pass the
/// same checks.
/// </remarks>
public sealed class PdfAFontSubstitution {
    private PdfAFontSubstitution(IReadOnlyList<PdfEmbeddedFontFamily> families, bool usesInstalledFonts) {
        Families = families;
        UsesInstalledFonts = usesInstalledFonts;
    }

    /// <summary>Embeds no substitutes: a source with a font it does not embed is refused.</summary>
    public static PdfAFontSubstitution None { get; } = new(Array.Empty<PdfEmbeddedFontFamily>(), usesInstalledFonts: false);

    /// <summary>Embeds substitutes from the fonts installed on this machine.</summary>
    public static PdfAFontSubstitution InstalledFonts { get; } = new(Array.Empty<PdfEmbeddedFontFamily>(), usesInstalledFonts: true);

    /// <summary>
    /// Embeds substitutes from the supplied families only. A family is matched by the name it was created with, against
    /// the font the PDF names, its standard-font substitutes (for example Arial or Liberation Sans for Helvetica), or
    /// <c>Symbol</c> and <c>ZapfDingbats</c>. Add <see cref="ThenInstalledFonts"/> to fall back to installed fonts.
    /// </summary>
    public static PdfAFontSubstitution Supplied(params PdfEmbeddedFontFamily[] families) {
        Guard.NotNull(families, nameof(families));
        if (families.Length == 0) throw new ArgumentException("Supply at least one font family.", nameof(families));
        if (families.Any(static family => family is null)) throw new ArgumentException("Font families cannot be null.", nameof(families));
        return new PdfAFontSubstitution(Array.AsReadOnly((PdfEmbeddedFontFamily[])families.Clone()), usesInstalledFonts: false);
    }

    /// <summary>Returns a substitution that tries these families first and then the fonts installed on this machine.</summary>
    public PdfAFontSubstitution ThenInstalledFonts() => UsesInstalledFonts ? this : new PdfAFontSubstitution(Families, usesInstalledFonts: true);

    /// <summary>Gets the supplied families, tried before installed fonts.</summary>
    public IReadOnlyList<PdfEmbeddedFontFamily> Families { get; }

    /// <summary>Gets whether fonts installed on this machine are tried.</summary>
    public bool UsesInstalledFonts { get; }

    internal bool IsOff => Families.Count == 0 && !UsesInstalledFonts;

    /// <summary>Yields every face for the candidate family names: supplied families first, then installed ones.</summary>
    internal IEnumerable<(string Family, string Origin, PdfEmbeddedFontFamily Font)> Resolve(IReadOnlyList<string> candidates) {
        foreach (string candidate in candidates) {
            string key = PdfEmbeddedFontFamily.NormalizeFamilyKey(candidate);
            foreach (PdfEmbeddedFontFamily family in Families) {
                if (string.Equals(PdfEmbeddedFontFamily.NormalizeFamilyKey(family.FamilyName), key, StringComparison.Ordinal)) {
                    yield return (family.FamilyName, "supplied", family);
                }
            }
        }
        if (!UsesInstalledFonts) yield break;
        foreach (string candidate in candidates) {
            if (PdfEmbeddedFontFamily.TryFromSystem(candidate, out PdfEmbeddedFontFamily? installed) && installed is not null) {
                yield return (candidate, "installed", installed);
            }
        }
    }
}
