namespace OfficeIMO.Pdf;

internal static class PdfFlowNestingRules {
    public static bool IsColumnFlowPrimitive(IPdfBlock block) =>
        block is HeadingBlock or RichParagraphBlock or BulletListBlock or NumberedListBlock or
        PanelParagraphBlock or TableBlock or HorizontalRuleBlock or ImageBlock or ShapeBlock or DrawingBlock or
        TextFieldBlock or CheckBoxBlock or ChoiceFieldBlock or RadioButtonGroupBlock or
        TextAnnotationBlock or FreeTextAnnotationBlock or HighlightAnnotationBlock or BookmarkBlock or SpacerBlock;
}
