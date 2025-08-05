using System.IO;

namespace OfficeIMO.Converters {
    public interface IWordConverter {
        void Convert(Stream input, Stream output, IConversionOptions options);
    }
}
