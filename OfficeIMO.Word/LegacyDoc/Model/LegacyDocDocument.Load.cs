using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal sealed partial class LegacyDocDocument {
        /// <summary>Loads a legacy DOC model while preserving terminal compound resource failures.</summary>
        /// <exception cref="OfficeInspectionResourceLimitException">The canonical compound reader reaches a configured admission ceiling.</exception>
        public static LegacyDocDocument Load(byte[] bytes, LegacyDocImportOptions? options = null) {
            try { return LoadCore(bytes, options); }
            catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("word", exception.LimitName, exception.MaximumValue, exception);
            }
        }
    }
}
