using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class OfficeLifecycleContractTests {
    [Fact]
    public void LifecycleOptionsUseSafeExplicitReadWriteDefaults() {
        var createOptions = new DocumentCreateOptions();
        var loadOptions = new DocumentLoadOptions();

        Assert.Equal(DocumentPersistenceMode.Explicit, createOptions.PersistenceMode);
        Assert.Equal(DocumentPersistenceMode.Explicit, loadOptions.PersistenceMode);
        Assert.Equal(DocumentAccessMode.ReadWrite, loadOptions.AccessMode);
    }
}
