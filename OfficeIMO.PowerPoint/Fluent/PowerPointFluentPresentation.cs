namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    ///     Provides a fluent API wrapper around <see cref="PowerPointPresentation" />.
    /// </summary>
    public class PowerPointFluentPresentation {
        /// <summary>
        ///     Initializes a new instance of the <see cref="PowerPointFluentPresentation" /> class.
        /// </summary>
        /// <param name="presentation">Presentation to wrap.</param>
        public PowerPointFluentPresentation(PowerPointPresentation presentation) {
            Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
        }

        internal PowerPointPresentation Presentation { get; }

        /// <summary>
        ///     Adds a new slide to the presentation.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public PowerPointSlideBuilder Slide(int masterIndex = 0, int layoutIndex = 0) {
            PowerPointSlide slide = Presentation.AddSlide(masterIndex, layoutIndex);
            return new PowerPointSlideBuilder(slide);
        }

        /// <summary>
        /// Ends fluent configuration and returns the underlying <see cref="PowerPointPresentation"/>.
        /// </summary>
        /// <returns>The wrapped <see cref="PowerPointPresentation"/> for further processing.</returns>
        public PowerPointPresentation End() {
            return Presentation;
        }
    }
}