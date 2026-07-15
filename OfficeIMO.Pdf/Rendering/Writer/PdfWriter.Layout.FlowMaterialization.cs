namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private readonly struct FlowMaterializationKey : System.IEquatable<FlowMaterializationKey> {
        private readonly FlowBlock _flow;
        private readonly int _pageNumber;
        private readonly double _availableHeight;
        private readonly double _fullContentHeight;
        private readonly double _contentWidth;
        private readonly double _pageWidth;
        private readonly double _pageHeight;

        public FlowMaterializationKey(FlowBlock flow, PdfFlowContext context) {
            _flow = flow;
            _pageNumber = context.PageNumber;
            _availableHeight = context.AvailableHeight;
            _fullContentHeight = context.FullContentHeight;
            _contentWidth = context.ContentWidth;
            _pageWidth = context.PageWidth;
            _pageHeight = context.PageHeight;
        }

        public bool Equals(FlowMaterializationKey other) =>
            ReferenceEquals(_flow, other._flow) &&
            _pageNumber == other._pageNumber &&
            _availableHeight.Equals(other._availableHeight) &&
            _fullContentHeight.Equals(other._fullContentHeight) &&
            _contentWidth.Equals(other._contentWidth) &&
            _pageWidth.Equals(other._pageWidth) &&
            _pageHeight.Equals(other._pageHeight);

        public override bool Equals(object? obj) => obj is FlowMaterializationKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_flow);
                hash = (hash * 397) ^ _pageNumber;
                hash = (hash * 397) ^ _availableHeight.GetHashCode();
                hash = (hash * 397) ^ _fullContentHeight.GetHashCode();
                hash = (hash * 397) ^ _contentWidth.GetHashCode();
                hash = (hash * 397) ^ _pageWidth.GetHashCode();
                hash = (hash * 397) ^ _pageHeight.GetHashCode();
                return hash;
            }
        }
    }
}
