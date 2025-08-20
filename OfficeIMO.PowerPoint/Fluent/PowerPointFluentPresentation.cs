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
        ///     Adds and optionally configures a new slide to the presentation.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        /// <param name="configure">Optional action used to configure the slide via a <see cref="PowerPointSlideBuilder"/>.</param>
        public PowerPointFluentPresentation Slide(int masterIndex = 0, int layoutIndex = 0, Action<PowerPointSlideBuilder> configure = null) {
            PowerPointSlide slide = Presentation.AddSlide(masterIndex, layoutIndex);
            if (configure != null) {
                PowerPointSlideBuilder builder = new PowerPointSlideBuilder(slide);
                configure(builder);
            }
            return this;
        }

        /// <summary>
        ///     Adds and configures a new slide using a builder action.
        /// </summary>
        /// <param name="configure">Action used to configure the slide.</param>
        public PowerPointFluentPresentation Slide(Action<PowerPointSlideBuilder> configure) {
            return Slide(0, 0, configure);
        }

        public PowerPointPresentation End() {
            return Presentation;
        }
    }
}