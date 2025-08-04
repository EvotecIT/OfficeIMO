using System.IO;

namespace OfficeIMO.Word.Converters {
    /// <summary>
    /// Represents generic conversion options.
    /// </summary>
    public interface IConversionOptions {
    }

    /// <summary>
    /// Defines a converter that transforms data using streams.
    /// </summary>
    public interface IWordConverter {
        /// <summary>
        /// Performs conversion using the provided streams and options.
        /// </summary>
        /// <param name="input">Source data stream.</param>
        /// <param name="output">Destination stream for converted data.</param>
        /// <param name="options">Conversion options controlling the operation.</param>
        void Convert(Stream input, Stream output, IConversionOptions options);
    }
}

