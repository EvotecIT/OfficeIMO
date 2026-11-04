using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel.LegacyXls.Model {
    internal sealed partial class LegacyXlsWorkbook {
        /// <summary>Loads a legacy XLS model while preserving terminal compound resource failures.</summary>
        /// <exception cref="OfficeInspectionResourceLimitException">The canonical compound reader reaches a configured admission ceiling.</exception>
        public static LegacyXlsWorkbook Load(byte[] bytes, LegacyXlsImportOptions? options = null) {
            try { return LoadCore(bytes, options); }
            catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("excel", exception.LimitName, exception.MaximumValue, exception);
            }
        }
    }
}
