namespace OfficeIMO.Visio {
    internal static class VisioSemanticUserCells {
        public const string Kind = "OfficeIMO.Kind";
        public const string CalloutKind = "Callout";
        public const string DiagramAdornmentKind = "DiagramAdornment";
        public const string BackgroundSurfaceKind = "BackgroundSurface";
        public const string DiagramAdornmentRole = "OfficeIMO.DiagramAdornmentRole";
        public const string GeneratedAdornmentRole = "Generated";
        public const string UserAdornmentRole = "User";
        public const string SequenceActivationKind = "SequenceActivation";
        public const string SequenceFragmentKind = "SequenceFragment";
        public const string SwimlaneLaneHeaderKind = "SwimlaneLaneHeader";
        public const string SwimlaneLaneKind = "SwimlaneLane";
        public const string SwimlanePhaseKind = "SwimlanePhase";
        public const string SwimlaneActivityKind = "SwimlaneActivity";
        public const string CalloutTargetId = "OfficeIMO.CalloutTargetId";
        public const string CalloutLeaderId = "OfficeIMO.CalloutLeaderId";
        public const string SwimlaneLaneId = "OfficeIMO.SwimlaneLaneId";
        public const string SwimlanePhaseId = "OfficeIMO.SwimlanePhaseId";
        public const string SwimlaneActivityType = "OfficeIMO.SwimlaneActivityType";
        public const string ContainerHeadingHeight = "OfficeIMO.ContainerHeadingHeight";
        public const string DataGraphicTargetId = "OfficeIMO.DataGraphicTargetId";
        public const string DataGraphicField = "OfficeIMO.DataGraphicField";
        public const string DataGraphicValue = "OfficeIMO.DataGraphicValue";
        public const string DataGraphicRole = "OfficeIMO.DataGraphicRole";
        public const string StencilId = "OfficeIMO.StencilId";
        public const string StencilName = "OfficeIMO.StencilName";
        public const string StencilCategory = "OfficeIMO.StencilCategory";
        public const string StencilCatalog = "OfficeIMO.StencilCatalog";
        public const string StencilSourcePackagePath = "OfficeIMO.StencilSourcePackagePath";
        public const string StencilKeywords = "OfficeIMO.StencilKeywords";
        public const string StencilAliases = "OfficeIMO.StencilAliases";
        public const string StencilTags = "OfficeIMO.StencilTags";
        public const string StencilIconNameU = "OfficeIMO.StencilIconNameU";
        public const string StencilDefaultWidth = "OfficeIMO.StencilDefaultWidth";
        public const string StencilDefaultHeight = "OfficeIMO.StencilDefaultHeight";
        public const string StencilDefaultUnit = "OfficeIMO.StencilDefaultUnit";
        public const string StencilPreviewImageRelationshipId = "OfficeIMO.StencilPreviewImageRelationshipId";
        public const string StencilPreviewImageTarget = "OfficeIMO.StencilPreviewImageTarget";
        public const string StencilPreviewImageContentType = "OfficeIMO.StencilPreviewImageContentType";
        public const string StencilPreviewImageExtension = "OfficeIMO.StencilPreviewImageExtension";
        public const string StencilPreviewImageByteLength = "OfficeIMO.StencilPreviewImageByteLength";

        public static void MarkGeneratedAdornment(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            shape.SetUserCell(Kind, DiagramAdornmentKind, "STR", prompt: "OfficeIMO semantic kind");
            shape.SetUserCell(DiagramAdornmentRole, GeneratedAdornmentRole, "STR", prompt: "OfficeIMO diagram adornment role");
        }

        public static bool IsGeneratedDiagramAdornment(VisioShape shape) {
            if (shape == null || !shape.IsDiagramAdornment) {
                return false;
            }

            return string.Equals(shape.GetUserCellValue(DiagramAdornmentRole), GeneratedAdornmentRole, StringComparison.OrdinalIgnoreCase);
        }
    }
}
