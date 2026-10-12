namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // Running stories are admitted before active containers resume. Their
        // continuation padding and bottom insets still reduce the body frame.
        private double GetRunningBodyContinuationReservation() {
            double height = 0D;
            foreach (ContainerRenderScope scope in activeContainerScopes)
                height += scope.Style.GetFragmentTopPadding(isContinuation: true) + scope.Style.FragmentBottomInset;
            return height;
        }

        private (double X, double Width, double ReservedHeight) GetRunningBodyContinuationFrame(PdfOptions source, int sequenceDepth) {
            double frameX = source.MarginLeft;
            double frameWidth = source.PageWidth - source.MarginLeft - source.MarginRight;
            double reserved = 0D;
            foreach (ContainerRenderScope scope in activeContainerScopes) {
                int ownerDepth = 0;
                while (ownerDepth < activeBlockSequences.Count &&
                    !ReferenceEquals(activeBlockSequences[ownerDepth].Owner, scope.Container)) ownerDepth++;
                reserved += scope.Style.GetFragmentTopPadding(isContinuation: true);
                if (ownerDepth <= sequenceDepth) {
                    var frame = ResolveContainerFrame(scope.Container, scope.Style, frameX, frameWidth);
                    frameX = frame.X + scope.Style.PaddingX;
                    frameWidth = frame.ContentWidth;
                    reserved += scope.Style.FragmentBottomInset;
                } else {
                    // An exhausted leaf resumes and closes before the first
                    // pending ancestor sibling; its width no longer applies.
                    reserved += scope.Style.BottomPadding + scope.Style.SpacingAfter;
                }
            }
            return (frameX, frameWidth, reserved);
        }
    }
}
