using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word;

internal sealed partial class WordMl2003Projection {
    private readonly Dictionary<int, HashSet<int>> _listLevels = new();

    /// <summary>Reconstructs declarations before styles and stories resolve their list references.</summary>
    private void ReadLists(XElement source) {
        CheckAttributes(source, "");
        var numbering = new Numbering();
        var definitions = new Dictionary<int, AbstractNum>();
        foreach (XElement child in source.Elements(Word + "listDef")) {
            _token.ThrowIfCancellationRequested();
            int id = ReadListIndex(child, "listDefId");
            if (definitions.ContainsKey(id)) throw new InvalidDataException("WordML 2003 contains a duplicate list definition identifier.");
            ValidateListLevels(child);
            var definition = NormalizeMappedProperties(new AbstractNum(MapProperties(child, "listDef")!.ToString(SaveOptions.DisableFormatting)));
            definition.AbstractNumberId = id;
            definitions.Add(id, definition);
            numbering.AppendChild(definition);
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
            if (!definitions.TryGetValue(definitionId, out AbstractNum? definition)) {
                Omit(child, "list instance with an unresolved definition");
                continue;
            }
            ValidateListLevels(child, overrides: true);
            var instance = NormalizeMappedProperties(new NumberingInstance(MapProperties(child, "list")!.ToString(SaveOptions.DisableFormatting)));
            instance.NumberID = id;
            instance.AbstractNumId = new AbstractNumId { Val = definitionId };
            var levels = new HashSet<int>(definition.Elements<Level>().Select(level => level.LevelIndex!.Value));
            foreach (LevelOverride item in instance.Elements<LevelOverride>()) {
                if (item.GetFirstChild<Level>() != null) levels.Add(item.LevelIndex!.Value);
                else if (!levels.Contains(item.LevelIndex!.Value))
                    throw new InvalidDataException("WordML 2003 list start override references an undefined level.");
            }
            _listLevels.Add(id, levels);
            numbering.AppendChild(instance);
        }
        if (numbering.HasChildren) {
            ReserveRelationship(_main, createsPart: true);
            _main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
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
            NumberFormatValues? format = code == 255 ? NumberFormatValues.None : LegacyDocNumberFormatMapper.FromNfc(code);
            if (!format.HasValue) return false;
            value = new EnumValue<NumberFormatValues>(format.Value).InnerText!;
            return true;
        }
        value = source.ToLowerInvariant() switch {
            "singlelevel" when context == "plt" => "singleLevel",
            "multilevel" when context == "plt" => "multilevel",
            "hybridmultilevel" when context == "plt" => "hybridMultilevel",
            "tab" or "space" or "nothing" when context == "suff" => source.ToLowerInvariant(),
            _ => source
        };
        return context == "plt" ? IsEnum<MultiLevelValues>(value) : IsEnum<LevelSuffixValues>(value);
    }
}
