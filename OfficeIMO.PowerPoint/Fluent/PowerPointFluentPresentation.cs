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
        ///     Adds and returns a builder for a new slide.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public PowerPointSlideBuilder Slide(int masterIndex = 0, int layoutIndex = 0) {
            PowerPointSlide slide = Presentation.AddSlide(masterIndex, layoutIndex);
            return new PowerPointSlideBuilder(this, slide);
        }

        /// <summary>
        ///     Duplicates a slide and returns a builder for the new slide.
        /// </summary>
        /// <param name="index">Index of the slide to duplicate.</param>
        /// <param name="insertAt">Index where the duplicate should be inserted. Defaults to index + 1.</param>
        public PowerPointSlideBuilder DuplicateSlide(int index, int? insertAt = null) {
            PowerPointSlide slide = Presentation.DuplicateSlide(index, insertAt);
            return new PowerPointSlideBuilder(this, slide);
        }

        /// <summary>
        ///     Imports a slide from another presentation and returns a builder for the new slide.
        /// </summary>
        /// <param name="sourcePresentation">Presentation to import from.</param>
        /// <param name="sourceIndex">Index of the slide to import.</param>
        /// <param name="insertAt">Index where the imported slide should be inserted. Defaults to end.</param>
        public PowerPointSlideBuilder ImportSlide(PowerPointPresentation sourcePresentation, int sourceIndex, int? insertAt = null) {
            PowerPointSlide slide = Presentation.ImportSlide(sourcePresentation, sourceIndex, insertAt);
            return new PowerPointSlideBuilder(this, slide);
        }

        /// <summary>
        ///     Duplicates a slide and applies configuration.
        /// </summary>
        /// <param name="index">Index of the slide to duplicate.</param>
        /// <param name="insertAt">Index where the duplicate should be inserted.</param>
        /// <param name="configure">Action used to configure the duplicated slide.</param>
        public PowerPointFluentPresentation DuplicateSlide(int index, int? insertAt, Action<PowerPointSlideBuilder> configure) {
            PowerPointSlideBuilder builder = DuplicateSlide(index, insertAt);
            configure?.Invoke(builder);
            return this;
        }

        /// <summary>
        ///     Imports a slide from another presentation and applies configuration.
        /// </summary>
        /// <param name="sourcePresentation">Presentation to import from.</param>
        /// <param name="sourceIndex">Index of the slide to import.</param>
        /// <param name="insertAt">Index where the imported slide should be inserted.</param>
        /// <param name="configure">Action used to configure the imported slide.</param>
        public PowerPointFluentPresentation ImportSlide(
            PowerPointPresentation sourcePresentation,
            int sourceIndex,
            int? insertAt,
            Action<PowerPointSlideBuilder> configure) {
            PowerPointSlideBuilder builder = ImportSlide(sourcePresentation, sourceIndex, insertAt);
            configure?.Invoke(builder);
            return this;
        }

        /// <summary>
        ///     Adds and optionally configures a new slide to the presentation.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        /// <param name="configure">Action used to configure the slide.</param>
        public PowerPointFluentPresentation Slide(int masterIndex, int layoutIndex, Action<PowerPointSlideBuilder> configure) {
            PowerPointSlideBuilder builder = Slide(masterIndex, layoutIndex);
            configure?.Invoke(builder);
            return this;
        }

        /// <summary>
        ///     Adds and configures a new slide using a builder action with default master and layout indexes.
        /// </summary>
        /// <param name="configure">Action used to configure the slide.</param>
        public PowerPointFluentPresentation Slide(Action<PowerPointSlideBuilder> configure) {
            return Slide(0, 0, configure);
        }

        /// <summary>
        ///     Completes fluent configuration and returns the underlying presentation.
        /// </summary>
        public PowerPointPresentation End() {
            return Presentation;
        }
    }
}
