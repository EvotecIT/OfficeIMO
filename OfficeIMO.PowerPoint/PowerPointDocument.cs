using System.Collections.Generic;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a PowerPoint presentation containing slides.
    /// </summary>
    public class PowerPointDocument {
        private readonly List<PowerPointSlide> _slides = new();

        /// <summary>
        /// Collection of slides in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSlide> Slides => _slides;

        /// <summary>
        /// Adds a new slide to the presentation.
        /// </summary>
        /// <param name="title">Title of the slide.</param>
        public PowerPointSlide AddSlide(string title) {
            PowerPointSlide slide = new(title);
            _slides.Add(slide);
            return slide;
        }
    }
}
