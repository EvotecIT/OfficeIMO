namespace OfficeIMO.Visio {
    public partial class VisioPage {
        /// <summary>
        /// Moves a swimlane activity to another lane/phase cell and relayouts swimlane activities.
        /// </summary>
        public VisioPage MoveSwimlaneActivity(string activityId, string laneId, string phaseId, Action<VisioSwimlaneRelayoutOptions> configureOptions) {
            ArgumentNullException.ThrowIfNull(configureOptions);

            VisioSwimlaneRelayoutOptions options = new();
            configureOptions(options);
            return this.MoveSwimlaneActivity(activityId, laneId, phaseId, options);
        }

    }
}
