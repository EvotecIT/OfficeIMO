namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a BIFF8 worksheet scenario parsed from a SCENARIO record.
    /// </summary>
    public sealed class LegacyXlsScenario {
        /// <summary>
        /// Creates worksheet scenario metadata.
        /// </summary>
        public LegacyXlsScenario(
            string name,
            bool locked,
            bool hidden,
            string? user,
            string? comment,
            IReadOnlyList<LegacyXlsScenarioInputCell> inputCells) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Locked = locked;
            Hidden = hidden;
            User = user;
            Comment = comment;
            InputCells = inputCells ?? throw new ArgumentNullException(nameof(inputCells));
        }

        /// <summary>Gets the scenario name.</summary>
        public string Name { get; }

        /// <summary>Gets whether the scenario is locked.</summary>
        public bool Locked { get; }

        /// <summary>Gets whether the scenario is hidden.</summary>
        public bool Hidden { get; }

        /// <summary>Gets the scenario owner name, when present.</summary>
        public string? User { get; }

        /// <summary>Gets the scenario comment, when present.</summary>
        public string? Comment { get; }

        /// <summary>Gets the changed cells and their stored values.</summary>
        public IReadOnlyList<LegacyXlsScenarioInputCell> InputCells { get; }
    }
}
