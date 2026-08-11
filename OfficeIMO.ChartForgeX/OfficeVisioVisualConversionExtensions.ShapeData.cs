using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using global::ChartForgeX.VisualArtifacts;

namespace OfficeIMO.ChartForgeX;

public static partial class OfficeVisioVisualConversionExtensions {
    private static void AddCommonShapeData(
        IDictionary<string, string?> target,
        string kind,
        string? status,
        string? groupId,
        IEnumerable<KeyValuePair<string, string>> metadata,
        OfficeVisioVisualConversionReport report,
        string context) {
        AddValue(target, "CFX.Kind", kind);
        AddValue(target, "CFX.Status", status);
        AddValue(target, "CFX.GroupId", groupId);
        AddMetadataData(target, "Metadata.", metadata, report, context);
    }

    private static void AddDetailData(
        IDictionary<string, string?> target,
        IReadOnlyList<VisualArtifactInterchangeDetail> details,
        OfficeVisioVisualConversionReport report,
        string context) {
        for (int index = 0; index < details.Count; index++) {
            VisualArtifactInterchangeDetail detail = details[index];
            string number = (index + 1).ToString(CultureInfo.InvariantCulture);
            string prefix = "Detail." + number + ".";
            string valueKey = prefix + detail.Label;
            if (IsReservedDetailField(detail.Label)) {
                valueKey = prefix + "Field." + detail.Label;
                report.Warn($"Detail label '{detail.Label}' on {context} detail {number} was projected as '{valueKey}' to avoid a reserved Visio Shape Data field collision.");
            }

            AddValue(target, valueKey, detail.Value);
            AddValue(target, prefix + "Label", detail.Label);
            AddValue(target, prefix + "Value", detail.Value);
            AddValue(target, prefix + "Icon", detail.IconId);
            AddValue(target, prefix + "Status", detail.Status);
            AddValue(target, prefix + "Color", detail.Color);
            AddMetadataData(target, prefix + "Metadata.", detail.Metadata, report, context + " detail " + number);
        }
    }

    private static bool IsReservedDetailField(string label) =>
        string.Equals(label, "Label", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(label, "Value", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(label, "Icon", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(label, "Status", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(label, "Color", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(label, "Metadata", StringComparison.OrdinalIgnoreCase) ||
        label.StartsWith("Metadata.", StringComparison.OrdinalIgnoreCase);

    private static void AddPortData(
        IDictionary<string, string?> target,
        IReadOnlyList<VisualArtifactInterchangePort> ports,
        OfficeVisioVisualConversionReport report,
        string context) {
        for (int index = 0; index < ports.Count; index++) {
            VisualArtifactInterchangePort port = ports[index];
            string number = (index + 1).ToString(CultureInfo.InvariantCulture);
            string prefix = "Port." + number + ".";
            AddValue(target, "Port." + number, port.Id + "|" + port.Side + "|" + port.Offset.ToString("R", CultureInfo.InvariantCulture));
            AddValue(target, prefix + "Id", port.Id);
            AddValue(target, prefix + "Side", port.Side);
            AddValue(target, prefix + "Offset", port.Offset.ToString("R", CultureInfo.InvariantCulture));
            AddValue(target, prefix + "Label", port.Label);
            AddMetadataData(target, prefix + "Metadata.", port.Metadata, report, context + " port " + number);
        }
    }

    private static void AddMetadataData(
        IDictionary<string, string?> target,
        string prefix,
        IEnumerable<KeyValuePair<string, string>> metadata,
        OfficeVisioVisualConversionReport report,
        string context) {
        foreach (KeyValuePair<string, string> item in metadata.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
            string requested = prefix + item.Key;
            string resolved = requested;
            int suffix = 2;
            while (target.Keys.Any(key => string.Equals(key, resolved, StringComparison.OrdinalIgnoreCase))) {
                resolved = requested + " [" + suffix.ToString(CultureInfo.InvariantCulture) + "]";
                suffix++;
            }
            if (!string.Equals(requested, resolved, StringComparison.Ordinal)) {
                report.Warn($"Metadata key '{item.Key}' on {context} was projected as '{resolved}' because Visio Shape Data names are case-insensitive.");
            }
            AddValue(target, resolved, item.Value);
        }
    }

    private static void AddValue(IDictionary<string, string?> target, string key, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) target[key] = value;
    }
}
