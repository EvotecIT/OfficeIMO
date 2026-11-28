using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private readonly Dictionary<int, HashSet<int>> _listLevels = new();

    /// <summary>Reconstructs declarations before styles and stories resolve their list references.</summary>
    private void ReadLists(XElement source) {
        CheckAttributes(source, "");
        OfficeXmlTree? numbering = null;
        var definitions = new Dictionary<int, HashSet<int>>();
        foreach (XElement child in source.Elements(Word + "listDef")) {
            _token.ThrowIfCancellationRequested();
            int id = ReadListIndex(child, "listDefId");
            if (definitions.ContainsKey(id)) throw new InvalidDataException("WordML 2003 contains a duplicate list definition identifier.");
            ValidateListLevels(child);
            XElement definition = MapProperties(child, "listDef")!;
            NormalizeMappedProperties(definition);
            definition.SetAttributeValue(ModernWord + "abstractNumId", id);
            definitions.Add(id, new HashSet<int>(definition.Elements(ModernWord + "lvl").Select(level => (int)level.Attribute(ModernWord + "ilvl")!)));
            numbering ??= AddMainPart("numbering", "numbering", out _);
            AppendMapped(numbering, numbering.Root, definition);
        }
        var ids = new HashSet<int>();
        foreach (XElement child in source.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "listDef") continue;
            if (child.Name != Word + "list") { Omit(child, "list-library content"); continue; }
            int id = ReadListIndex(child, "ilfo", positive: true);
            if (!ids.Add(id)) throw new InvalidDataException("WordML 2003 contains a duplicate list instance identifier.");
            XElement? reference = child.Element(Word + "ilst");
            if (reference == null || reference.ElementsAfterSelf(Word + "ilst").Any())
                throw new InvalidDataException("WordML 2003 list must reference exactly one list definition.");
            int definitionId = ReadListIndex(reference, "val");
            if (!definitions.TryGetValue(definitionId, out HashSet<int>? definition)) {
                Omit(child, "list instance with an unresolved definition");
                continue;
            }
            ValidateListLevels(child, overrides: true);
            XElement instance = MapProperties(child, "list")!;
            NormalizeMappedProperties(instance);
            instance.SetAttributeValue(ModernWord + "numId", id);
            instance.Element(ModernWord + "abstractNumId")?.Remove();
            instance.AddFirst(new XElement(ModernWord + "abstractNumId", new XAttribute(ModernWord + "val", definitionId)));
            var levels = new HashSet<int>(definition);
            foreach (XElement item in instance.Elements(ModernWord + "lvlOverride")) {
                int index = (int)item.Attribute(ModernWord + "ilvl")!;
                if (item.Element(ModernWord + "lvl") != null) levels.Add(index);
                else if (!levels.Contains(index))
                    throw new InvalidDataException("WordML 2003 list start override references an undefined level.");
            }
            _listLevels.Add(id, levels);
            AppendMapped(numbering!, numbering!.Root, instance);
        }
    }

    private void ValidateListLevels(XElement container, bool overrides = false) {
        var levels = new HashSet<int>();
        foreach (XElement level in container.Elements(Word + (overrides ? "lvlOverride" : "lvl"))) {
            _token.ThrowIfCancellationRequested();
            int index = ReadListIndex(level, "ilvl", maximum: 8);
            if (!levels.Add(index)) throw new InvalidDataException("WordML 2003 contains a duplicate list level.");
            if (!overrides) continue;
            if (level.Elements(Word + "startOverride").Skip(1).Any() || level.Elements(Word + "lvl").Skip(1).Any())
                throw new InvalidDataException("WordML 2003 contains duplicate list override content.");
            if (level.Element(Word + "lvl") is { } replacement && ReadListIndex(replacement, "ilvl", maximum: 8) != index)
                throw new InvalidDataException("WordML 2003 list override and replacement level do not match.");
        }
    }

    private XElement? MapListProperties(XElement source) {
        CheckAttributes(source, "");
        XElement? instance = null;
        XElement? level = null;
        foreach (XElement child in source.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name == Word + "ilfo" || child.Name == Word + "ilvl") {
                if (child.Name == Word + "ilfo") {
                    if (instance != null) throw new InvalidDataException("WordML 2003 contains a duplicate paragraph list instance.");
                    instance = child;
                } else {
                    if (level != null) throw new InvalidDataException("WordML 2003 contains a duplicate paragraph list level.");
                    level = child;
                }
                CheckAttributes(child, "val");
                foreach (XElement nested in child.Elements()) Omit(nested, "nested list-reference content");
            } else if (child.Name == Auxiliary + "t" || child.Name == Auxiliary + "font") {
                // Cached display hints are superseded by the editable numbering definition.
                CheckAttributes(child, "val", Auxiliary);
                foreach (XElement nested in child.Elements()) Omit(nested, "nested list display hint");
            } else Omit(child, "list properties");
        }
        int? numberId = instance == null ? null : ReadListIndex(instance, "val");
        int? levelId = level == null ? null : ReadListIndex(level, "val", maximum: 8);
        if (numberId > 0 && (!_listLevels.TryGetValue(numberId.Value, out HashSet<int>? declared) || !declared.Contains(levelId ?? 0))) {
            Omit(source, "paragraph list with an unresolved instance or level");
            return null;
        }
        var output = new XElement(ModernWord + "numPr");
        output.SetAttributeValue(XNamespace.Xmlns + "w", ModernWord.NamespaceName);
        if (levelId.HasValue) output.Add(new XElement(ModernWord + "ilvl", new XAttribute(ModernWord + "val", levelId.Value)));
        if (numberId.HasValue) output.Add(new XElement(ModernWord + "numId", new XAttribute(ModernWord + "val", numberId.Value)));
        return output;
    }

    private static int ReadListIndex(XElement source, string attribute, bool positive = false, int maximum = int.MaxValue) {
        if (!int.TryParse((string?)source.Attribute(Word + attribute), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ||
            value < (positive ? 1 : 0) || value > maximum)
            throw new InvalidDataException("WordML 2003 contains an invalid " + source.Name.LocalName + " list identifier or level.");
        return value;
    }

    private static bool TryMapListValue(string context, string source, out string value) {
        value = source;
        if (context == "nfc") {
            if (!byte.TryParse(source, NumberStyles.Integer, CultureInfo.InvariantCulture, out byte code)) return false;
            LegacyDocNumberFormat? format = code == 255 ? LegacyDocNumberFormat.None : LegacyDocNumberFormatMapper.FromNfc(code);
            if (!format.HasValue) return false;
            value = format.Value.ToXml();
            return true;
        }
        value = source.ToLowerInvariant() switch {
            "singlelevel" when context == "plt" => "singleLevel",
            "multilevel" when context == "plt" => "multilevel",
            "hybridmultilevel" when context == "plt" => "hybridMultilevel",
            "tab" or "space" or "nothing" when context == "suff" => source.ToLowerInvariant(),
            _ => source
        };
        return context == "plt" ? ContainsName("singleLevel multilevel hybridMultilevel", value)
            : WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(value, out WordListLevelSuffix _);
    }
}
