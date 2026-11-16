namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>
        /// Lists the blocks a keep chain can reach from <paramref name="startIndex"/>: transparent wrappers
        /// are measured through their children, and the end of a transparent wrapper continues in its parent.
        /// </summary>
        private IEnumerable<IPdfBlock> EnumerateFollowingKeepBlocks(IList<IPdfBlock> blocks, int startIndex) {
            int scopeIndex = activeBlockSequences.FindLastIndex(scope => ReferenceEquals(scope.Blocks, blocks));
            while (true) {
                for (int index = startIndex; index < blocks.Count; index++) {
                    foreach (IPdfBlock block in ExpandTransparentMeasurementBlocks(new[] { blocks[index] })) yield return block;
                }
                if (scopeIndex <= 0 || GetTransparentMeasurementChildren(activeBlockSequences[scopeIndex].Owner) == null) yield break;
                BlockSequenceScope parent = activeBlockSequences[--scopeIndex];
                blocks = parent.Blocks;
                startIndex = parent.Index + 1;
            }
        }
    }
}
