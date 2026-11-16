namespace OfficeIMO.Pdf;

/// <summary>
/// The PDF/UA-1 condition catalogue used by <see cref="PdfDocumentProof.CheckAccessibility"/>: every Matterhorn
/// Protocol 1.1 failure condition plus the ISO 14289-1 requirements that veraPDF checks without a Matterhorn number.
/// </summary>
public static class PdfAccessibilityConditions {
    private const PdfAccessibilityCoverage A = PdfAccessibilityCoverage.Automatic;
    private const PdfAccessibilityCoverage H = PdfAccessibilityCoverage.ManualReview;
    private const PdfAccessibilityCoverage N = PdfAccessibilityCoverage.NotChecked;

    private static readonly PdfAccessibilityCondition[] _all = {
        C("01", "01-001", "Artifact content is tagged as real content.", "7.1", H),
        C("01", "01-002", "Real content is marked as an artifact.", "7.1", H),
        C("01", "01-003", "Content marked as Artifact is inside tagged content.", "7.1", A, V("7.1-1")),
        C("01", "01-004", "Tagged content is inside content marked as Artifact.", "7.1", A, V("7.1-2")),
        C("01", "01-005", "Content is neither marked as Artifact nor tagged as real content.", "7.1", A, V("7.1-3"), V("8.2.2-1")),
        C("01", "01-006", "A structure element's type or attributes do not fit its content.", "7.1", H),
        C("01", "01-007", "The MarkInfo Suspects entry is true.", "7.1", A, V("7.1-4")),
        C("01", "UA1:6.2-1", "The catalog has no MarkInfo dictionary with Marked set to true.", "6.2", A, V("6.2-1"), V("6.2-1")),
        C("01", "UA1:7.1-11", "The catalog has no StructTreeRoot, so the document has no logical structure.", "7.1", A, V("7.1-11"), V("8.2.1-1")),
        C("01", "UA1:7.1-12", "A structure element has no P (parent) entry.", "7.1", A, V("7.1-12"), V("8.2.1-2")),

        C("02", "02-001", "A non-standard structure type is not role mapped to a standard type.", "7.1", A, V("7.1-5"), V("8.2.4-1")),
        C("02", "02-002", "A non-standard structure type is mapped to a semantically inappropriate standard type.", "7.1", H),
        C("02", "02-003", "The role map contains a circular mapping.", "7.1", A, V("7.1-6"), V("8.2.4-2")),
        C("02", "02-004", "A standard structure type is remapped.", "7.1", A, V("7.1-7"), V("8.2.4-3")),

        C("03", "03-001", "An action causes flickering.", "7.1", H),
        C("03", "03-002", "Multimedia content flickers.", "7.1", H),
        C("03", "03-003", "JavaScript causes flickering.", "7.1", H),
        C("04", "04-001", "Meaning carried by colour, contrast, format or layout is not reflected in the tags.", "7.1", H),
        C("05", "05-001", "A media annotation has audio that is not available in another form.", "7.1", H),
        C("05", "05-002", "An audio annotation's content is not available in another form.", "7.1", H),
        C("05", "05-003", "JavaScript beeps without another means of notification.", "7.1", H),

        C("06", "06-001", "The catalog has no XMP metadata stream.", "7.1", A, V("7.1-8"), V("8.11.1-2")),
        C("06", "06-002", "The XMP metadata has no PDF/UA identification (pdfuaid:part).", "5", A, V("5-1", "5-2"), V("5-1", "5-2")),
        C("06", "06-003", "The XMP metadata has no dc:title.", "7.1", A, V("7.1-9"), V("8.11.1-1")),
        C("06", "06-004", "dc:title does not clearly identify the document.", "7.1", H),

        C("07", "07-001", "The ViewerPreferences dictionary has no DisplayDocTitle entry.", "7.1", A, V("7.1-10"), V("8.11.2-1")),
        C("07", "07-002", "ViewerPreferences DisplayDocTitle is false.", "7.1", A, V("7.1-10"), V("8.11.2-1")),

        C("08", "08-001", "OCR text contains significant errors.", "7.1", H),
        C("08", "08-002", "OCR text is not tagged.", "7.1", H),

        C("09", "09-001", "Tags are not in logical reading order.", "7.2", H),
        C("09", "09-002", "Structure elements are nested in a semantically inappropriate way.", "7.2", H),
        C("09", "09-003", "A structure type does not fit its content after role mapping.", "7.2", H),
        C("09", "09-004", "A table structure element breaks the ISO 32000-1 Table 337 nesting rules.", "7.2", A,
            V("7.2-3", "7.2-4", "7.2-5", "7.2-6", "7.2-7", "7.2-8", "7.2-9", "7.2-10", "7.2-11", "7.2-12", "7.2-13", "7.2-14", "7.2-16", "7.2-36", "7.2-37", "7.2-38", "7.2-39")),
        C("09", "09-005", "A list structure element breaks the ISO 32000-1 Table 336 nesting rules.", "7.2", A, V("7.2-17", "7.2-18", "7.2-19", "7.2-20", "7.2-40")),
        C("09", "09-006", "A table-of-contents structure element breaks the ISO 32000-1 Table 333 nesting rules.", "7.2", A, V("7.2-26", "7.2-27", "7.2-28")),
        C("09", "09-007", "A Ruby structure element breaks the ISO 32000-1 Table 338 nesting rules.", "7.2", A, V(), V("8.2.5.23-1")),
        C("09", "09-008", "A Warichu structure element breaks the ISO 32000-1 Table 338 nesting rules.", "7.2", A, V(), V("8.2.5.24-1")),

        C("10", "10-001", "A character code cannot be mapped to Unicode.", "7.2", N, V("7.21.7-1"), V("8.4.5.8-1")),

        C("11", "11-001", "The natural language of page content text cannot be determined.", "7.2", A, V("7.2-34"), V("8.4.4-1")),
        C("11", "11-002", "The natural language of Alt, ActualText or E text cannot be determined.", "7.2", A, V("7.2-21", "7.2-22", "7.2-23", "7.2-30", "7.2-31", "7.2-32"), V("8.4.4-1")),
        C("11", "11-003", "The natural language of outline entries cannot be determined.", "7.2", A, V("7.2-2"), V("8.4.4-1")),
        C("11", "11-004", "The natural language of annotation Contents cannot be determined.", "7.2", A, V("7.2-24"), V("8.4.4-1")),
        C("11", "11-005", "The natural language of a form field's TU entry cannot be determined.", "7.2", A, V("7.2-25"), V("8.4.4-1")),
        C("11", "11-006", "The natural language of document metadata cannot be determined.", "7.2", A, V("7.2-33"), V("8.4.4-1")),
        C("11", "11-007", "The declared natural language is not appropriate.", "7.2", H),
        C("11", "UA1:7.2-29", "A Lang value is not a valid language identifier.", "7.2", A, V("7.2-29"), V("8.4.4-2")),

        C("12", "12-001", "Stretched characters are not represented appropriately.", "7.2", H),

        C("13", "13-001", "Graphics other than text and artifacts are not tagged as Figure.", "7.3", H),
        C("13", "13-002", "A link with a meaningful graphic lacks alternative text for both.", "7.3", H),
        C("13", "13-003", "A caption is not tagged as Caption.", "7.3", H),
        C("13", "13-004", "A Figure has no alternative text or replacement text.", "7.3", A, V("7.3-1"), V("8.2.5.28.2-1")),
        C("13", "13-005", "ActualText is used on a Figure where alternative text fits better.", "7.3", H),
        C("13", "13-006", "A graphic that only has meaning as part of a group is tagged on its own.", "7.3", H),
        C("13", "13-007", "A more accessible representation is available but not used.", "7.3", H),
        C("13", "13-008", "A Figure meant to be read as text has no ActualText.", "7.3", H),

        C("14", "14-001", "Headings are not tagged.", "7.4", H),
        C("14", "14-002", "The document uses numbered headings but the first one is not H1.", "7.4", A, V("7.4.2-1")),
        C("14", "14-003", "Numbered heading levels are skipped going deeper.", "7.4", A, V("7.4.2-1")),
        C("14", "14-004", "Numbered heading tags are not role mapped to Arabic-numbered heading types.", "7.4", H),
        C("14", "14-005", "A seventh or deeper heading level does not use an H7 or deeper tag.", "7.4", H),
        C("14", "14-006", "A structure node has more than one H child.", "7.4", A, V("7.4.4-1")),
        C("14", "14-007", "The document uses both H and numbered H# tags.", "7.4", A, V("7.4.4-2", "7.4.4-3"), V("8.2.5.12-1")),

        C("15", "15-001", "A row header cell is not tagged as a header.", "7.5", H),
        C("15", "15-002", "A column header cell is not tagged as a header.", "7.5", H),
        C("15", "15-003", "In a table without Headers and IDs, a TH cell has no Scope attribute.", "7.5", A, V("7.5-1", "7.5-2"), V("8.2.5.26-5", "8.2.5.26-6")),
        C("15", "15-004", "Content that is not tabular is tagged as a table.", "7.5", H),
        C("15", "15-005", "A cell's header cannot be determined unambiguously.", "7.5", H),

        C("16", "16-001", "An ordered list has no ListNumbering attribute.", "7.6", H),
        C("16", "16-002", "An ordered list's ListNumbering value is not a numbering style.", "7.6", H),
        C("16", "16-003", "A list is not tagged as a list.", "7.6", H),

        C("17", "17-001", "A mathematical expression is not tagged as Formula.", "7.7", H),
        C("17", "17-002", "A Formula has no Alt or ActualText.", "7.7", A, V("7.7-1")),
        C("17", "17-003", "Formula text does not meet the Unicode mapping requirements.", "7.7", N, V("7.21.7-1")),

        C("18", "18-001", "Headers and footers are not marked as pagination artifacts.", "7.8", H),
        C("18", "18-002", "Header and footer artifacts are not given the Header or Footer subtype.", "7.8", H),

        C("19", "19-001", "Footnotes or endnotes are not tagged as Note.", "7.9", H),
        C("19", "19-002", "References are not tagged as Reference.", "7.9", H),
        C("19", "19-003", "A Note has no ID entry.", "7.9", A, V("7.9-1")),
        C("19", "19-004", "A Note's ID is not unique.", "7.9", A, V("7.9-2")),

        C("20", "20-001", "An optional content configuration in Configs has no Name or an empty Name.", "7.10", A, V("7.10-1"), V("8.7-1")),
        C("20", "20-002", "The default optional content configuration (D) has no Name or an empty Name.", "7.10", A, V("7.10-1"), V("8.7-1")),
        C("20", "20-003", "An optional content configuration contains an AS entry.", "7.10", A, V("7.10-2"), V("8.7-2")),

        C("21", "21-001", "An embedded file's file specification lacks F or UF.", "7.11", A, V("7.11-1")),
        C("22", "22-001", "Article threads do not follow the logical reading order.", "7.12", H),
        C("23", "23-001", "Digital signatures have no specific test; their form fields follow the form rules.", "7.13", H),
        C("24", "24-001", "A non-interactive form is not tagged with the PrintField attribute.", "7.14", H),
        C("25", "25-001", "The XFA configuration sets dynamicRender to required.", "7.15", A, V("7.15-1"), V("8.10.1-3")),
        C("26", "26-001", "The file is encrypted but the encryption dictionary has no P entry.", "7.16", A, V("7.16-1")),
        C("26", "26-002", "The file is encrypted and the P entry denies content access for accessibility (bit 10 is off).", "7.16", A, V("7.16-1")),
        C("27", "27-001", "Navigation has no specific test; it relies on correct semantics.", "7.17", H),

        C("28", "28-001", "An annotation is not in the correct reading order.", "7.18", H),
        C("28", "28-002", "An annotation other than Widget, Link or PrinterMark is not inside an Annot structure element.", "7.18.1", A, V("7.18.1-1"), V("8.9.2.3-1")),
        C("28", "28-003", "An annotation is used for visual formatting without matching tags.", "7.18.1", H),
        C("28", "28-004", "An annotation other than Widget has no Contents and its structure element has no Alt.", "7.18.1", A, V("7.18.1-2")),
        C("28", "28-005", "A form field has no TU and its structure element has no Alt.", "7.18.1", A, V("7.18.1-3"), V("8.10.2.3-1")),
        C("28", "28-006", "An annotation of a subtype not defined in ISO 32000 does not meet the annotation rules.", "7.18.2", A, V("7.18.1-1", "7.18.1-2")),
        C("28", "28-007", "A TrapNet annotation is present.", "7.18.2", A, V("7.18.2-1"), V("8.9.2.4.15-1")),
        C("28", "28-008", "A page with annotations has no Tabs entry.", "7.18.3", A, V("7.18.3-1"), V("8.9.3.3-1")),
        C("28", "28-009", "A page with annotations has a Tabs value other than S.", "7.18.3", A, V("7.18.3-1")),
        C("28", "28-010", "A Widget annotation is not inside a Form structure element.", "7.18.4", A, V("7.18.4-1"), V("8.10.1-1")),
        C("28", "UA1:7.18.4-2", "A Form structure element without a Role attribute has more than its single widget reference.", "7.18.4", A, V("7.18.4-2"), V("8.10.1-2")),
        C("28", "28-011", "A Link annotation is not inside a Link structure element.", "7.18.5", A, V("7.18.5-1"), V("8.2.5.20-1")),
        C("28", "28-012", "A Link annotation has no Contents alternate description.", "7.18.5", A, V("7.18.5-2")),
        C("28", "28-013", "IsMap is true and the function is not offered another way.", "7.18.5", H),
        C("28", "28-014", "A media clip data dictionary has no CT entry.", "7.18.6", A, V("7.18.6.2-1")),
        C("28", "28-015", "A media clip data dictionary has no Alt entry.", "7.18.6", A, V("7.18.6.2-2")),
        C("28", "28-016", "A FileAttachment annotation's file specification lacks F or UF.", "7.18.7", A, V("7.11-1"), V("8.9.2.4.10-1")),
        C("28", "28-017", "A PrinterMark annotation is part of the logical structure.", "7.18.8", A, V("7.18.8-1"), V("8.9.2.4.14-1")),
        C("28", "28-018", "A PrinterMark annotation's appearance is not marked as Artifact.", "7.18.8", N),
        C("29", "29-001", "A script needs keystrokes with specific timing.", "7.19", H),

        C("30", "30-001", "A reference XObject is present.", "7.20", A, V("7.20-1")),
        C("30", "30-002", "A Form XObject with MCIDs is painted more than once.", "7.20", A, V("7.20-2")),

        C("31", "31-001", "A Type 0 font with a non-Identity CMap has different CIDSystemInfo Registry values.", "7.21.3.1", A, V("7.21.3.1-1"), V("8.4.5.3.1-1")),
        C("31", "31-002", "A Type 0 font with a non-Identity CMap has different CIDSystemInfo Ordering values.", "7.21.3.1", A, V("7.21.3.1-1"), V("8.4.5.3.1-1")),
        C("31", "31-003", "A Type 0 font with a non-Identity CMap has incompatible CIDSystemInfo Supplement values.", "7.21.3.1", A, V("7.21.3.1-1"), V("8.4.5.3.1-1")),
        C("31", "31-004", "A Type 2 CIDFont's CIDToGIDMap is neither a stream nor Identity.", "7.21.3.2", A, V("7.21.3.2-1"), V("8.4.5.3.2-1")),
        C("31", "31-005", "A Type 2 CIDFont has no CIDToGIDMap entry.", "7.21.3.2", A, V("7.21.3.2-1"), V("8.4.5.3.2-1")),
        C("31", "31-006", "A CMap is neither predefined nor embedded.", "7.21.3.3", A, V("7.21.3.3-1"), V("8.4.5.4-1")),
        C("31", "31-007", "An embedded CMap's WMode differs from the CMap dictionary's WMode.", "7.21.3.3", A, V("7.21.3.3-2"), V("8.4.5.4-2")),
        C("31", "31-008", "A CMap references another CMap that is not predefined.", "7.21.3.3", A, V("7.21.3.3-3"), V("8.4.5.4-3")),
        C("31", "31-009", "A font used to render text is not embedded.", "7.21.4.1", A, V("7.21.4.1-1"), V("8.4.5.5.1-1")),
        C("31", "31-010", "An embedded font program is not legally embeddable for unlimited rendering.", "7.21.4.1", H),
        C("31", "31-011", "An embedded font program lacks glyphs referenced by rendered text.", "7.21.4.1", N, V("7.21.4.1-2"), V("8.4.5.5.1-2")),
        C("31", "31-012", "A Type 1 CharSet does not list every glyph in the font program.", "7.21.4.2", N, V("7.21.4.2-1")),
        C("31", "31-013", "A Type 1 CharSet lists glyphs missing from the font program.", "7.21.4.2", N, V("7.21.4.2-1")),
        C("31", "31-014", "A CIDSet does not identify every CID in the font program.", "7.21.4.2", N, V("7.21.4.2-2")),
        C("31", "31-015", "A CIDSet identifies CIDs missing from the font program.", "7.21.4.2", N, V("7.21.4.2-2")),
        C("31", "31-016", "Glyph widths in the font dictionary and the font program differ.", "7.21.5", N, V("7.21.5-1"), V("8.4.5.6-1")),
        C("31", "31-017", "A non-symbolic TrueType font program has no non-symbolic cmap.", "7.21.6", N, V("7.21.6-1"), V("8.4.5.7-1")),
        C("31", "31-018", "A rendered glyph of a non-symbolic TrueType font cannot be found through its cmap.", "7.21.6", N, V("7.21.6-1"), V("8.4.5.7-1")),
        C("31", "31-019", "A non-symbolic TrueType font has no Encoding entry.", "7.21.6", A, V("7.21.6-2"), V("8.4.5.7-2")),
        C("31", "31-020", "A non-symbolic TrueType font's Encoding dictionary has no BaseEncoding.", "7.21.6", A, V("7.21.6-2"), V("8.4.5.7-2")),
        C("31", "31-021", "A non-symbolic TrueType font's encoding is neither MacRomanEncoding nor WinAnsiEncoding.", "7.21.6", A, V("7.21.6-2"), V("8.4.5.7-2")),
        C("31", "31-022", "A non-symbolic TrueType Differences array uses names outside the Adobe Glyph List.", "7.21.6", N, V("7.21.6-2")),
        C("31", "31-023", "A non-symbolic TrueType font with Differences has no (3,1) cmap.", "7.21.6", N, V("7.21.6-1")),
        C("31", "31-024", "A symbolic TrueType font has an Encoding entry.", "7.21.6", A, V("7.21.6-3"), V("8.4.5.7-3")),
        C("31", "31-025", "A symbolic TrueType font program has no cmap.", "7.21.6", N, V("7.21.6-4"), V("8.4.5.7-4")),
        C("31", "31-026", "A symbolic TrueType font program has several cmaps but no (3,0) cmap.", "7.21.6", N, V("7.21.6-4"), V("8.4.5.7-4")),
        C("31", "31-027", "A font has no ToUnicode map and its character codes cannot be mapped to Unicode another way.", "7.21.7", A, V("7.21.7-1"), V("8.4.5.8-1")),
        C("31", "31-028", "A ToUnicode map contains the value 0.", "7.21.7", N, V("7.21.7-2"), V("8.4.5.8-2")),
        C("31", "31-029", "A ToUnicode map contains U+FEFF or U+FFFE.", "7.21.7", N, V("7.21.7-2"), V("8.4.5.8-2")),
        C("31", "31-030", "Text references the .notdef glyph.", "7.21.8", N, V("7.21.8-1"), V("8.4.5.9-1")),
    };

    private static readonly Dictionary<string, PdfAccessibilityCondition> _byId =
        _all.ToDictionary(static condition => condition.ConditionId, StringComparer.Ordinal);

    /// <summary>Every condition, in Matterhorn checkpoint order.</summary>
    public static IReadOnlyList<PdfAccessibilityCondition> All { get; } = Array.AsReadOnly(_all);

    /// <summary>Returns the condition with this identifier.</summary>
    public static PdfAccessibilityCondition Get(string conditionId) {
        Guard.NotNull(conditionId, nameof(conditionId));
        return _byId.TryGetValue(conditionId, out PdfAccessibilityCondition? condition)
            ? condition
            : throw new ArgumentOutOfRangeException(nameof(conditionId), conditionId, "Unknown accessibility condition.");
    }

    private static string[] V(params string[] rules) => rules;

    private static PdfAccessibilityCondition C(string checkpoint, string id, string description, string clause,
        PdfAccessibilityCoverage coverage, string[]? veraPdf = null, string[]? pdfUa2 = null) =>
        new PdfAccessibilityCondition(checkpoint, id, description, clause, coverage, veraPdf ?? Array.Empty<string>(), pdfUa2 ?? Array.Empty<string>());
}
