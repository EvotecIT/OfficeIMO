using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Converters {
    public interface IWordConverter {
        void Convert(Stream input, Stream output, IConversionOptions options);
        Task ConvertAsync(Stream input, Stream output, IConversionOptions options);
    }
}
