using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

internal static partial class StudioProcessProbe {
    private static int VerifyLongTempIpc(string root) {
        try {
            Assert.True(Path.GetTempPath().Length > 104);
            string name = StudioSingleInstance.CreateName(root);
            using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
            Assert.True(primary.IsPrimary);
            bool accepted = false;
            primary.StartListening(_ => { accepted = true; return true; });
            using StudioSingleInstance later = Task.Factory.StartNew(
                () => StudioSingleInstance.Acquire(name), TaskCreationOptions.LongRunning).GetAwaiter().GetResult();
            Assert.False(later.IsPrimary);
            Assert.False(later.TryForward([Path.Combine(root, "source.pdf")], TimeSpan.FromMilliseconds(200)));
            Assert.False(accepted);
            Console.WriteLine("VERIFIED independent launch");
            return 0;
        } catch (Exception error) {
            Console.Error.WriteLine(error);
            return 1;
        }
    }
}
