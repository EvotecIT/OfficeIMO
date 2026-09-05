namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderSemanticBlock(SemanticBlock semantic) {
            flowSemanticScopes.Add(new FlowSemanticScope(semantic.Role, semantic.AlternativeText));
            try {
                ProcessBlocks(semantic.Blocks);
            } finally {
                flowSemanticScopes.RemoveAt(flowSemanticScopes.Count - 1);
            }
        }

        private static string MapSemanticStructureType(PdfSemanticRole role) => role switch {
            PdfSemanticRole.Part => "Part",
            PdfSemanticRole.Article => "Art",
            PdfSemanticRole.Section => "Sect",
            PdfSemanticRole.Division => "Div",
            PdfSemanticRole.BlockQuote => "BlockQuote",
            PdfSemanticRole.Caption => "Caption",
            PdfSemanticRole.Figure => "Figure",
            PdfSemanticRole.Form => "Form",
            _ => throw new ArgumentOutOfRangeException(nameof(role))
        };

        private sealed class FlowSemanticScope {
            public FlowSemanticScope(PdfSemanticRole role, string? alternativeText) {
                Role = role;
                AlternativeText = alternativeText;
            }

            public PdfSemanticRole Role { get; }
            public string? AlternativeText { get; }
            public PageStructElement? Element { get; set; }
            public LayoutResult.Page? ElementPage { get; set; }
        }
    }
}
