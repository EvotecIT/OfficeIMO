using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a single slide within a PowerPoint presentation.
    /// </summary>
    public class PowerPointSlide {
        private readonly List<PowerPointShape> _shapes = new();

        public PowerPointSlide(string title) {
            Title = title;
        }

        /// <summary>
        /// Title of the slide.
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// Shapes placed on the slide.
        /// </summary>
        public IList<PowerPointShape> Shapes => _shapes;
    }
}
