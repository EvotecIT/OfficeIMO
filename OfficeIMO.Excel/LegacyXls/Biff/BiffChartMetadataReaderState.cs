namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Tracks shallow chart-record nesting while scanning a chart substream.
    /// </summary>
    internal sealed class BiffChartMetadataReaderState {
        private int _depth;
        private int _sequence;

        internal BiffChartNestingInfo Capture(ushort recordType) {
            _sequence++;
            int depthBefore = _depth;
            string transition;

            if (recordType == 0x1033) {
                _depth++;
                transition = "Begin";
            } else if (recordType == 0x1034) {
                if (_depth == 0) {
                    transition = "UnmatchedEnd";
                } else {
                    _depth--;
                    transition = "End";
                }
            } else {
                transition = _depth == 0 ? "OutsideContainer" : "InsideContainer";
            }

            return new BiffChartNestingInfo(_sequence, depthBefore, _depth, transition);
        }
    }

    internal readonly struct BiffChartNestingInfo {
        internal BiffChartNestingInfo(int sequenceIndex, int depthBefore, int depthAfter, string transition) {
            SequenceIndex = sequenceIndex;
            DepthBefore = depthBefore;
            DepthAfter = depthAfter;
            Transition = transition;
        }

        internal int SequenceIndex { get; }

        internal int DepthBefore { get; }

        internal int DepthAfter { get; }

        internal string Transition { get; }
    }
}
