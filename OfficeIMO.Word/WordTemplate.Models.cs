using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>Controls how plain-text Word template placeholders are bound.</summary>
    public sealed class WordTemplateOptions {
        private CultureInfo _culture = CultureInfo.InvariantCulture;

        /// <summary>Gets or sets the culture used to format scalar values. The default is invariant culture.</summary>
        public CultureInfo Culture {
            get => _culture;
            set => _culture = value ?? throw new ArgumentNullException(nameof(value));
        }

        /// <summary>Gets or sets whether unresolved scalar placeholders are removed. The default preserves them for inspection.</summary>
        public bool RemoveMissingPlaceholders { get; set; }
    }

    /// <summary>Describes an embedded image supplied to a Word template placeholder.</summary>
    public sealed class WordTemplateImage {
        /// <summary>Creates an embedded image value from encoded image bytes.</summary>
        /// <param name="content">Encoded image bytes, such as PNG or JPEG content.</param>
        /// <param name="fileName">File name with an extension that identifies the image format.</param>
        /// <param name="width">Optional width in pixels.</param>
        /// <param name="height">Optional height in pixels.</param>
        /// <param name="description">Alternative-text description stored with the image.</param>
        public WordTemplateImage(byte[] content, string fileName, double? width = null, double? height = null, string description = "") {
            if (content == null) throw new ArgumentNullException(nameof(content));
            if (content.Length == 0) throw new ArgumentException("Image content cannot be empty.", nameof(content));
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("Image file name cannot be empty.", nameof(fileName));
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width), "Image width must be positive.");
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height), "Image height must be positive.");

            _content = (byte[])content.Clone();
            FileName = fileName;
            Width = width;
            Height = height;
            Description = description ?? string.Empty;
        }

        /// <summary>Gets a copy of the encoded image bytes.</summary>
        public byte[] Content => (byte[])_content.Clone();

        private readonly byte[] _content;

        /// <summary>Gets the image file name used to identify its format.</summary>
        public string FileName { get; }

        /// <summary>Gets the optional image width in pixels.</summary>
        public double? Width { get; }

        /// <summary>Gets the optional image height in pixels.</summary>
        public double? Height { get; }

        /// <summary>Gets the image alternative-text description.</summary>
        public string Description { get; }

        internal byte[] GetContentUnsafe() => _content;
    }

    /// <summary>Describes an external hyperlink supplied to a Word template placeholder.</summary>
    public sealed class WordTemplateHyperlink {
        /// <summary>Creates a hyperlink value.</summary>
        public WordTemplateHyperlink(string text, Uri uri, string tooltip = "", bool addStyle = true) {
            if (text == null) throw new ArgumentNullException(nameof(text));
            Uri = uri ?? throw new ArgumentNullException(nameof(uri));
            if (!uri.IsAbsoluteUri) throw new ArgumentException("Template hyperlinks require an absolute URI.", nameof(uri));
            Text = text;
            Tooltip = tooltip ?? string.Empty;
            AddStyle = addStyle;
        }

        /// <summary>Gets the visible hyperlink text.</summary>
        public string Text { get; }

        /// <summary>Gets the absolute hyperlink target.</summary>
        public Uri Uri { get; }

        /// <summary>Gets the optional hyperlink tooltip.</summary>
        public string Tooltip { get; }

        /// <summary>Gets whether the standard Word hyperlink style is applied.</summary>
        public bool AddStyle { get; }
    }

    /// <summary>Summarizes one Word template binding pass.</summary>
    public sealed class WordTemplateResult {
        internal WordTemplateResult(
            int placeholderCount,
            int replacedPlaceholderCount,
            int repeatedBlockCount,
            int conditionalBlockCount,
            IEnumerable<string> missingValueNames) {
            PlaceholderCount = placeholderCount;
            ReplacedPlaceholderCount = replacedPlaceholderCount;
            RepeatedBlockCount = repeatedBlockCount;
            ConditionalBlockCount = conditionalBlockCount;
            MissingValueNames = new ReadOnlyCollection<string>(missingValueNames
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(static name => name, StringComparer.OrdinalIgnoreCase)
                .ToArray());
        }

        /// <summary>Gets the number of scalar or rich placeholders discovered.</summary>
        public int PlaceholderCount { get; }

        /// <summary>Gets the number of scalar or rich placeholders replaced.</summary>
        public int ReplacedPlaceholderCount { get; }

        /// <summary>Gets the number of repeated block instances generated.</summary>
        public int RepeatedBlockCount { get; }

        /// <summary>Gets the number of conditional blocks evaluated.</summary>
        public int ConditionalBlockCount { get; }

        /// <summary>Gets unresolved scalar or rich placeholder names.</summary>
        public IReadOnlyList<string> MissingValueNames { get; }

        /// <summary>Gets whether every scalar or rich placeholder had a supplied value.</summary>
        public bool IsComplete => MissingValueNames.Count == 0;

        /// <summary>Throws when unresolved placeholders remain; otherwise returns this result.</summary>
        public WordTemplateResult EnsureComplete() {
            if (!IsComplete) {
                throw new InvalidOperationException("Word template values were not supplied for: " + string.Join(", ", MissingValueNames) + ".");
            }
            return this;
        }
    }
}
