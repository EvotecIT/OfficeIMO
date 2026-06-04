using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Moves a swimlane activity to another lane/phase cell and relayouts swimlane activities.
        /// </summary>
        public VisioFluentPage MoveSwimlaneActivity(string activityId, string laneId, string phaseId, VisioSwimlaneRelayoutOptions? options = null) {
            Page.MoveSwimlaneActivity(activityId, laneId, phaseId, options);
            return this;
        }

        /// <summary>
        /// Moves a swimlane activity to another lane/phase cell using inline relayout options.
        /// </summary>
        public VisioFluentPage MoveSwimlaneActivity(string activityId, string laneId, string phaseId, Action<VisioSwimlaneRelayoutOptions> configureOptions) {
            if (configureOptions == null) {
                throw new ArgumentNullException(nameof(configureOptions));
            }

            VisioSwimlaneRelayoutOptions options = new();
            configureOptions(options);
            return MoveSwimlaneActivity(activityId, laneId, phaseId, options);
        }

        /// <summary>
        /// Re-centers and stacks swimlane activities inside their current lane/phase cells.
        /// </summary>
        public VisioFluentPage RelayoutSwimlanes(VisioSwimlaneRelayoutOptions? options = null) {
            Page.RelayoutSwimlaneActivities(options);
            return this;
        }

        /// <summary>
        /// Re-centers and stacks swimlane activities using inline relayout options.
        /// </summary>
        public VisioFluentPage RelayoutSwimlanes(Action<VisioSwimlaneRelayoutOptions> configureOptions) {
            if (configureOptions == null) {
                throw new ArgumentNullException(nameof(configureOptions));
            }

            VisioSwimlaneRelayoutOptions options = new();
            configureOptions(options);
            return RelayoutSwimlanes(options);
        }

        /// <summary>
        /// Inspects discovered swimlane activity placements on the current page.
        /// </summary>
        public VisioFluentPage SwimlaneActivities(Action<IReadOnlyList<VisioSwimlaneActivityPlacement>> inspect) {
            if (inspect == null) {
                throw new ArgumentNullException(nameof(inspect));
            }

            inspect(Page.GetSwimlaneActivities());
            return this;
        }
    }
}
