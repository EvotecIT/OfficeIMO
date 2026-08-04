using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Visio {
    internal static class VisioPackageFormat {
        internal const string DrawingContentType = "application/vnd.ms-visio.drawing.main+xml";
        internal const string TemplateContentType = "application/vnd.ms-visio.template.main+xml";
        internal const string StencilContentType = "application/vnd.ms-visio.stencil.main+xml";
        internal const string MacroDrawingContentType = "application/vnd.ms-visio.drawing.macroEnabled.main+xml";
        internal const string MacroTemplateContentType = "application/vnd.ms-visio.template.macroEnabled.main+xml";
        internal const string MacroStencilContentType = "application/vnd.ms-visio.stencil.macroEnabled.main+xml";

        private static readonly IReadOnlyDictionary<string, VisioPackageType> ByContentType =
            new Dictionary<string, VisioPackageType>(StringComparer.OrdinalIgnoreCase) {
                [DrawingContentType] = VisioPackageType.Drawing,
                [TemplateContentType] = VisioPackageType.Template,
                [StencilContentType] = VisioPackageType.Stencil,
                [MacroDrawingContentType] = VisioPackageType.MacroEnabledDrawing,
                [MacroTemplateContentType] = VisioPackageType.MacroEnabledTemplate,
                [MacroStencilContentType] = VisioPackageType.MacroEnabledStencil
            };

        internal static bool TryFromContentType(string? contentType, out VisioPackageType type) {
            type = VisioPackageType.Drawing;
            return contentType != null && ByContentType.TryGetValue(contentType, out type);
        }

        internal static string GetContentType(VisioPackageType type) => type switch {
            VisioPackageType.Drawing => DrawingContentType,
            VisioPackageType.Template => TemplateContentType,
            VisioPackageType.Stencil => StencilContentType,
            VisioPackageType.MacroEnabledDrawing => MacroDrawingContentType,
            VisioPackageType.MacroEnabledTemplate => MacroTemplateContentType,
            VisioPackageType.MacroEnabledStencil => MacroStencilContentType,
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };

        internal static VisioPackageType FromPath(string path) =>
            Path.GetExtension(path).ToLowerInvariant() switch {
                ".vstx" => VisioPackageType.Template,
                ".vssx" => VisioPackageType.Stencil,
                ".vsdm" => VisioPackageType.MacroEnabledDrawing,
                ".vstm" => VisioPackageType.MacroEnabledTemplate,
                ".vssm" => VisioPackageType.MacroEnabledStencil,
                _ => VisioPackageType.Drawing
            };

        internal static bool IsMacroEnabled(VisioPackageType type) =>
            type == VisioPackageType.MacroEnabledDrawing ||
            type == VisioPackageType.MacroEnabledTemplate ||
            type == VisioPackageType.MacroEnabledStencil;

        internal static bool IsTemplate(VisioPackageType type) =>
            type == VisioPackageType.Template || type == VisioPackageType.MacroEnabledTemplate;

        internal static bool IsStencil(VisioPackageType type) =>
            type == VisioPackageType.Stencil || type == VisioPackageType.MacroEnabledStencil;
    }
}
