using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        /// <summary>Translates compound quotas from package loading and nested VBA/OLE projections before malformed fallbacks.</summary>
        internal static LegacyPptPresentation Load(byte[] bytes, LegacyPptImportOptions? options, CancellationToken cancellationToken) {
            try { return LoadCore(bytes, options, cancellationToken); }
            catch (OfficeCompoundStreamLimitExceededException exception) {
                throw new OfficeInspectionResourceLimitException("powerpoint", exception.LimitName, exception.MaximumValue, exception);
            }
        }
    }
}
