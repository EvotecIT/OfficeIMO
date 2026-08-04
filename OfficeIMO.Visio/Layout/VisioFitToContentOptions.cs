namespace OfficeIMO.Visio {
    /// <summary>Advanced resize-to-content behavior for imported pages.</summary>
    public sealed class VisioFitToContentOptions {
        /// <summary>Horizontal page margin in inches.</summary>
        public double HorizontalMargin { get; set; } = 0.5D;
        /// <summary>Vertical page margin in inches.</summary>
        public double VerticalMargin { get; set; } = 0.5D;
        /// <summary>Whether grouped child geometry contributes to bounds.</summary>
        public bool IncludeGroupChildren { get; set; } = true;
        /// <summary>Whether connector routes and label boxes contribute to bounds.</summary>
        public bool IncludeConnectors { get; set; } = true;
        /// <summary>Whether content is translated to the requested margins.</summary>
        public bool MoveContent { get; set; } = true;
        /// <summary>Whether page dimensions are resized around content.</summary>
        public bool ResizePage { get; set; } = true;
    }
}
