using System;
using System.Collections.Generic;
using global::ChartForgeX.Topology;
using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.ChartForgeX;

static partial class OfficeVisioVisualConversionExtensions {
    private static OfficeStrokeDashStyle? MapGraphLineStyle(TopologyEdgeLineStyle? lineStyle) {
        return lineStyle switch {
            TopologyEdgeLineStyle.Solid => OfficeStrokeDashStyle.Solid,
            TopologyEdgeLineStyle.Dashed => OfficeStrokeDashStyle.Dash,
            TopologyEdgeLineStyle.Dotted => OfficeStrokeDashStyle.Dot,
            _ => null
        };
    }

    private static OfficeStrokeDashStyle MapSequenceLineStyle(SequenceArtifactMessageLineStyle lineStyle) {
        return lineStyle == SequenceArtifactMessageLineStyle.Dashed
            ? OfficeStrokeDashStyle.Dash
            : OfficeStrokeDashStyle.Solid;
    }

    private static Color? MapNativeColor(
        string? value,
        string entityKind,
        string entityId,
        OfficeVisioVisualConversionReport report) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (Color.TryParseCss(value, out Color color)) {
            if (color.A == byte.MaxValue) return color;
            report.Warn(
                OfficeVisioVisualDiagnosticCode.ColorNotProjected,
                ParseEntityKind(entityKind),
                entityId,
                "colorAlpha",
                $"{entityKind} '{entityId}' color '{value}' remains in the CFX envelope and, when enabled, Shape Data because native Visio color projection does not currently preserve CSS alpha.");
            return null;
        }
        report.Warn(
            OfficeVisioVisualDiagnosticCode.ColorNotProjected,
            ParseEntityKind(entityKind),
            entityId,
            "color",
            $"{entityKind} '{entityId}' color '{value}' remains in the CFX envelope and, when enabled, Shape Data because it is not a supported native Visio color token.");
        return null;
    }

    private static void PreserveTooltipFidelity(
        IDictionary<string, string?> shapeData,
        string? tooltip,
        string? href,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        string entityKind,
        string entityId) {
        if (string.IsNullOrWhiteSpace(tooltip) || options.IncludeHyperlinks && !string.IsNullOrWhiteSpace(href)) return;
        if (options.IncludeShapeData) {
            AddValue(shapeData, "CFX.Tooltip", tooltip);
            string reason = options.IncludeHyperlinks
                ? "native hyperlink descriptions require a hyperlink address"
                : "native hyperlink projection was disabled";
            report.Warn(
                OfficeVisioVisualDiagnosticCode.TooltipRetainedAsShapeData,
                ParseEntityKind(entityKind),
                entityId,
                "tooltip",
                $"{entityKind} '{entityId}' tooltip was retained as Shape Data because {reason}.");
            return;
        }

        report.Warn(
            OfficeVisioVisualDiagnosticCode.TooltipNotProjected,
            ParseEntityKind(entityKind),
            entityId,
            "tooltip",
            $"{entityKind} '{entityId}' tooltip remains only in the CFX envelope because it could not be attached to a native hyperlink and Shape Data projection was disabled.");
    }

    private static void PreserveHyperlinkFidelity(
        IDictionary<string, string?> shapeData,
        string? href,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report,
        string entityKind,
        string entityId) {
        if (string.IsNullOrWhiteSpace(href) || options.IncludeHyperlinks) return;
        if (options.IncludeShapeData) AddValue(shapeData, "CFX.Href", href);
        report.Warn(
            OfficeVisioVisualDiagnosticCode.HyperlinkNotProjected,
            ParseEntityKind(entityKind),
            entityId,
            "href",
            options.IncludeShapeData
                ? $"{entityKind} '{entityId}' hyperlink was retained as Shape Data but was not projected as an active native hyperlink because hyperlink projection was disabled."
                : $"{entityKind} '{entityId}' hyperlink remains only in the CFX envelope because hyperlink and Shape Data projection were disabled.");
    }

    private static void ReportDisabledShapeDataFidelity(
        VisualArtifactInterchangeEnvelope envelope,
        OfficeVisioVisualOptions options,
        OfficeVisioVisualConversionReport report) {
        if (options.IncludeShapeData || envelope.Groups.Count + envelope.Nodes.Count + envelope.Edges.Count + envelope.Annotations.Count == 0) return;
        report.Warn(
            OfficeVisioVisualDiagnosticCode.ShapeDataDisabled,
            OfficeVisioVisualEntityKind.Artifact,
            envelope.Id,
            "shapeData",
            "Native entity Shape Data projection was disabled; source roles, free-form kinds, metrics, extensions, details, ports, and other non-visual semantics remain only in the CFX envelope.");
    }

    private static OfficeVisioVisualEntityKind ParseEntityKind(string entityKind) {
        if (entityKind.IndexOf("participant", StringComparison.OrdinalIgnoreCase) >= 0) return OfficeVisioVisualEntityKind.Participant;
        if (entityKind.IndexOf("message", StringComparison.OrdinalIgnoreCase) >= 0) return OfficeVisioVisualEntityKind.Message;
        if (entityKind.IndexOf("group", StringComparison.OrdinalIgnoreCase) >= 0) return OfficeVisioVisualEntityKind.Group;
        if (entityKind.IndexOf("edge", StringComparison.OrdinalIgnoreCase) >= 0) return OfficeVisioVisualEntityKind.Edge;
        return OfficeVisioVisualEntityKind.Node;
    }
}
