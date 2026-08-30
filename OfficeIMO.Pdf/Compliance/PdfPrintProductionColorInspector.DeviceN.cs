namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionColorInspector {
    private static bool HasValidDeviceNAttributes(
        PdfDictionary attributes,
        HashSet<string> deviceNColorants,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int maximumDecodedStreamBytes,
        ColorSpaceAliases? aliases,
        bool normalizeInlineImageAbbreviations,
        HashSet<PdfArray> activeArrays,
        int depth) {
        if (!HasOnlyKeys(attributes, "Subtype", "Colorants", "Process", "MixingHints")) return false;

        string subtype = "DeviceN";
        if (attributes.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)) {
            if (ResolveObject(objects, subtypeObject, depth, maximumObjectDepth) is not PdfName subtypeName ||
                subtypeName.Name is not ("DeviceN" or "NChannel")) return false;
            subtype = subtypeName.Name;
        }

        var describedSpotColorants = new HashSet<string>(StringComparer.Ordinal);
        if (attributes.Items.TryGetValue("Colorants", out PdfObject? colorantsObject)) {
            if (ResolveObject(objects, colorantsObject, depth, maximumObjectDepth) is not PdfDictionary colorants) return false;
            foreach (KeyValuePair<string, PdfObject> entry in colorants.Items) {
                if (ResolveObject(objects, entry.Value, depth + 1, maximumObjectDepth) is not PdfArray separation ||
                    separation.Items.Count != 4 ||
                    ResolveObject(objects, separation.Items[0], depth + 1, maximumObjectDepth) is not PdfName family ||
                    !string.Equals(family.Name, "Separation", StringComparison.Ordinal) ||
                    ResolveObject(objects, separation.Items[1], depth + 1, maximumObjectDepth) is not PdfName separationName ||
                    !string.Equals(separationName.Name, entry.Key, StringComparison.Ordinal)) return false;
                ColorSpaceUsage usage = ClassifyColorSpaceCore(
                    separation,
                    objects,
                    maximumObjectDepth,
                    maximumDecodedStreamBytes,
                    aliases,
                    normalizeInlineImageAbbreviations,
                    activeArrays,
                    depth + 1);
                if (!usage.IsKnown || usage.UsesPattern || usage.ComponentCount != 1) return false;
                describedSpotColorants.Add(entry.Key);
            }
        }

        var processComponents = new HashSet<string>(StringComparer.Ordinal);
        if (attributes.Items.TryGetValue("Process", out PdfObject? processObject)) {
            if (ResolveObject(objects, processObject, depth, maximumObjectDepth) is not PdfDictionary process ||
                !HasOnlyKeys(process, "ColorSpace", "Components") ||
                !process.Items.TryGetValue("ColorSpace", out PdfObject? processColorSpace) ||
                !process.Items.TryGetValue("Components", out PdfObject? componentsObject) ||
                !IsDeviceOrCieProcessColorSpace(processColorSpace, objects, maximumObjectDepth, depth + 1) ||
                ResolveObject(objects, componentsObject, depth + 1, maximumObjectDepth) is not PdfArray components ||
                !TryReadUniqueNames(components, objects, maximumObjectDepth, depth + 1, out processComponents)) {
                return false;
            }
            ColorSpaceUsage processUsage = ClassifyColorSpaceCore(
                processColorSpace,
                objects,
                maximumObjectDepth,
                maximumDecodedStreamBytes,
                aliases,
                normalizeInlineImageAbbreviations,
                activeArrays,
                depth + 1);
            if (!processUsage.IsKnown || processUsage.UsesPattern ||
                processUsage.ComponentCount != components.Items.Count) return false;
        }

        if (string.Equals(subtype, "NChannel", StringComparison.Ordinal)) {
            foreach (string colorant in deviceNColorants) {
                if (!processComponents.Contains(colorant) && !describedSpotColorants.Contains(colorant)) return false;
            }
        }

        if (attributes.Items.TryGetValue("MixingHints", out PdfObject? mixingObject)) {
            if (ResolveObject(objects, mixingObject, depth, maximumObjectDepth) is not PdfDictionary mixingHints ||
                !HasOnlyKeys(mixingHints, "Solidities", "PrintingOrder")) return false;
            bool hasSolidities = mixingHints.Items.TryGetValue("Solidities", out PdfObject? soliditiesObject);
            bool hasPrintingOrder = mixingHints.Items.TryGetValue("PrintingOrder", out PdfObject? printingOrderObject);
            if (hasSolidities) {
                if (!hasPrintingOrder ||
                    ResolveObject(objects, soliditiesObject!, depth + 1, maximumObjectDepth) is not PdfDictionary solidities) return false;
                foreach (PdfObject solidityObject in solidities.Items.Values) {
                    if (ResolveObject(objects, solidityObject, depth + 1, maximumObjectDepth) is not PdfNumber solidity ||
                        double.IsNaN(solidity.Value) || double.IsInfinity(solidity.Value) ||
                        solidity.Value < 0D || solidity.Value > 1D) return false;
                }
            }
            if (hasPrintingOrder) {
                if (ResolveObject(objects, printingOrderObject!, depth + 1, maximumObjectDepth) is not PdfArray printingOrder ||
                    !TryReadUniqueNames(printingOrder, objects, maximumObjectDepth, depth + 1, out HashSet<string> orderedColorants)) {
                    return false;
                }
                foreach (string colorant in deviceNColorants) {
                    if (!orderedColorants.Contains(colorant)) return false;
                }
            }
        }

        return true;
    }

    private static bool IsDeviceOrCieProcessColorSpace(
        PdfObject value,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int depth) {
        PdfObject? resolved = ResolveObject(objects, value, depth, maximumObjectDepth);
        if (resolved is PdfName name) {
            return name.Name is "DeviceGray" or "DeviceRGB" or "DeviceCMYK";
        }
        if (resolved is not PdfArray { Items.Count: > 0 } array ||
            ResolveObject(objects, array.Items[0], depth + 1, maximumObjectDepth) is not PdfName family) return false;
        return family.Name is "CalGray" or "CalRGB" or "Lab" or "ICCBased";
    }

    private static bool HasOnlyKeys(PdfDictionary dictionary, params string[] allowedKeys) {
        foreach (string key in dictionary.Items.Keys) {
            bool allowed = false;
            for (int index = 0; index < allowedKeys.Length; index++) {
                if (string.Equals(key, allowedKeys[index], StringComparison.Ordinal)) {
                    allowed = true;
                    break;
                }
            }
            if (!allowed) return false;
        }
        return true;
    }

    private static bool TryReadUniqueNames(
        PdfArray values,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth,
        int depth,
        out HashSet<string> names) {
        names = new HashSet<string>(StringComparer.Ordinal);
        if (values.Items.Count < 1) return false;
        for (int index = 0; index < values.Items.Count; index++) {
            if (ResolveObject(objects, values.Items[index], depth, maximumObjectDepth) is not PdfName name ||
                !names.Add(name.Name)) return false;
        }
        return true;
    }
}
