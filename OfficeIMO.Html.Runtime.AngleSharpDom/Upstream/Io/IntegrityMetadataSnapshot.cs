namespace AngleSharp.Io
{
    /// <summary>
    /// Carries the integrity selection made for one prepared resource request.
    /// </summary>
    public sealed class IntegrityMetadataSnapshot
    {
        private readonly object _sync = new object();
        private string? _value;
        private bool _isResolved;

        /// <summary>Creates a snapshot from the element attribute state at preparation.</summary>
        public IntegrityMetadataSnapshot(bool hasElementMetadata, string? value)
        {
            _value = value;
            _isResolved = hasElementMetadata;
        }

        /// <summary>Gets whether fallback selection has completed.</summary>
        public bool IsResolved
        {
            get { lock (_sync) return _isResolved; }
        }

        /// <summary>Gets the selected metadata after resolution.</summary>
        public string? Value
        {
            get { lock (_sync) return _value; }
        }

        /// <summary>Selects fallback metadata once when no element attribute supplied the value.</summary>
        public string? Resolve(string? fallback)
        {
            lock (_sync)
            {
                if (!_isResolved)
                {
                    _value = fallback;
                    _isResolved = true;
                }

                return _value;
            }
        }
    }
}
