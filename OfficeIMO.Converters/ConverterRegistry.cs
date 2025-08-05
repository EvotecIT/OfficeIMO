namespace OfficeIMO.Converters {
    /// <summary>
    /// Provides registration and resolution of <see cref="IWordConverter"/> instances by key.
    /// </summary>
    public static class ConverterRegistry {
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, System.Func<IWordConverter>> _converters = new(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers a converter factory under the specified key.
        /// </summary>
        /// <param name="key">Key describing the conversion, e.g. "markdown->word".</param>
        /// <param name="factory">Factory delegate creating a converter instance.</param>
        public static void Register(string key, System.Func<IWordConverter> factory) {
            if (string.IsNullOrWhiteSpace(key)) {
                throw new System.ArgumentException("Key cannot be null or whitespace.", nameof(key));
            }
            if (factory == null) {
                throw new System.ArgumentNullException(nameof(factory));
            }
            _converters[key] = factory;
        }

        /// <summary>
        /// Resolves a converter registered under the given key.
        /// </summary>
        /// <param name="key">Key describing the conversion.</param>
        /// <returns>An instance of <see cref="IWordConverter"/>.</returns>
        public static IWordConverter Resolve(string key) {
            if (key == null) {
                throw new System.ArgumentNullException(nameof(key));
            }
            if (_converters.TryGetValue(key, out var factory)) {
                return factory();
            }
            throw new System.InvalidOperationException($"Converter for key '{key}' is not registered.");
        }
    }
}
