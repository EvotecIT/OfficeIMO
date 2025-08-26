using System;

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