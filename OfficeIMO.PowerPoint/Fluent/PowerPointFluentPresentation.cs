namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    /// Provides a fluent API wrapper around <see cref="PowerPointPresentation"/>.
    /// </summary>
    public class PowerPointFluentPresentation {
        internal PowerPointPresentation Presentation { get; }

        public PowerPointFluentPresentation(PowerPointPresentation presentation) {
            Presentation = presentation ?? throw new System.ArgumentNullException(nameof(presentation));
        }

        /// <summary>
        /// Adds a new slide to the presentation.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public SlideBuilder Slide(int masterIndex = 0, int layoutIndex = 0) {
            PowerPointSlide slide = Presentation.AddSlide(masterIndex, layoutIndex);
            return new SlideBuilder(slide);
        }
    }
}