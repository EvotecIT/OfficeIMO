using OfficeIMO.Security;

namespace OfficeIMO.Tests {
    public partial class Word {
        private static IOfficeSecurityProvider SecurityProvider => OfficeSecurityProvider.Default;
    }
}
