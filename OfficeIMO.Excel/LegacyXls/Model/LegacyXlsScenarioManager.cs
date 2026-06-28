namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents BIFF8 scenario-manager metadata parsed from a ScenMan record.
    /// </summary>
    public sealed class LegacyXlsScenarioManager {
        /// <summary>
        /// Creates scenario-manager metadata.
        /// </summary>
        public LegacyXlsScenarioManager(ushort scenarioCount, short currentScenarioIndex, short shownScenarioIndex, IReadOnlyList<string> resultRanges) {
            ScenarioCount = scenarioCount;
            CurrentScenarioIndex = currentScenarioIndex;
            ShownScenarioIndex = shownScenarioIndex;
            ResultRanges = resultRanges ?? throw new ArgumentNullException(nameof(resultRanges));
        }

        /// <summary>Gets the scenario count declared by ScenMan.</summary>
        public ushort ScenarioCount { get; }

        /// <summary>Gets the current scenario index, or -1 when no scenario is current.</summary>
        public short CurrentScenarioIndex { get; }

        /// <summary>Gets the shown scenario index, or -1 when no scenario is shown.</summary>
        public short ShownScenarioIndex { get; }

        /// <summary>Gets result-cell ranges declared by ScenMan.</summary>
        public IReadOnlyList<string> ResultRanges { get; }
    }
}
