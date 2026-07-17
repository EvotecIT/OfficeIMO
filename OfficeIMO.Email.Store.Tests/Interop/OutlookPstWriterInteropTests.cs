using OfficeIMO.Email;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeIMO.Email.Store.Tests;

public sealed class OutlookPstWriterInteropTests {
    [OutlookInteropFact]
    public void Generated_unicode_pst_can_be_mounted_read_and_removed_by_classic_outlook() {
        string? retainedPath = Environment.GetEnvironmentVariable(
            "OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP_OUTPUT");
        string path = string.IsNullOrWhiteSpace(retainedPath)
            ? Path.Combine(Path.GetTempPath(),
                string.Concat("officeimo-outlook-interop-", Guid.NewGuid().ToString("N"), ".pst"))
            : Path.GetFullPath(retainedPath!);
        Exception? failure = null;
        var thread = new Thread(() => {
            try { RunInterop(path, !string.IsNullOrWhiteSpace(retainedPath)); }
            catch (Exception exception) { failure = exception; }
        }) { IsBackground = true, Name = "OfficeIMO Outlook PST interoperability" };
#pragma warning disable CA1416
        thread.SetApartmentState(ApartmentState.STA);
#pragma warning restore CA1416
        thread.Start();
        bool completed = thread.Join(TimeSpan.FromMinutes(2));
        if (!completed && string.IsNullOrWhiteSpace(retainedPath)) TryDelete(path);
        Assert.True(completed,
            "Classic Outlook interoperability did not finish within two minutes.");
        if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static void RunInterop(string path, bool retainOutput) {
#pragma warning disable CA1416
        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
#pragma warning restore CA1416
        Assert.NotNull(outlookType);

        object? application = null;
        object? nameSpace = null;
        object? stores = null;
        object? store = null;
        object? root = null;
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions("OfficeIMO Interop"))) {
                if (!string.Equals(Environment.GetEnvironmentVariable(
                    "OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP_EMPTY"), "1", StringComparison.Ordinal)) {
                    string folder = writer.AddFolder("OfficeIMO Synthetic");
                    writer.AddItem(folder, new EmailDocument {
                        Subject = "OfficeIMO synthetic interoperability item",
                        MessageClass = "IPM.Note"
                    });
                }
                writer.Complete();
            }
            if (retainOutput) {
                File.Copy(path, string.Concat(path, ".before-outlook"), overwrite: true);
            }
            application = Activator.CreateInstance(outlookType);
            Assert.NotNull(application);
            dynamic outlook = application!;
            nameSpace = outlook.GetNamespace("MAPI");
            dynamic mapi = nameSpace!;
            mapi.AddStoreEx(path, 2); // OlStoreType.olStoreUnicode
            stores = mapi.Stores;
            dynamic outlookStores = stores!;
            for (int index = 1; index <= outlookStores.Count; index++) {
                dynamic candidate = outlookStores.Item(index);
                if (string.Equals(Convert.ToString(candidate.FilePath), path,
                    StringComparison.OrdinalIgnoreCase)) {
                    store = candidate;
                    break;
                }
                Release(candidate);
            }
            Assert.NotNull(store);
            dynamic mountedStore = store!;
            root = mountedStore.GetRootFolder();
            dynamic mountedRoot = root!;
            if (!string.Equals(Environment.GetEnvironmentVariable(
                "OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP_EMPTY"), "1", StringComparison.Ordinal)) {
                dynamic folderObject = mountedRoot.Folders.Item("OfficeIMO Synthetic");
                Assert.Equal(1, (int)folderObject.Items.Count);
                dynamic item = folderObject.Items.Item(1);
                Assert.Equal("OfficeIMO synthetic interoperability item",
                    Convert.ToString(item.Subject));
                Release(item);
                Release(folderObject);
            }
            mapi.RemoveStore(mountedRoot);
            Release(root);
            root = null;
        } finally {
            if (root != null && nameSpace != null) {
                try { ((dynamic)nameSpace).RemoveStore((dynamic)root); }
                catch (COMException) { }
            }
            Release(root);
            Release(store);
            Release(stores);
            Release(nameSpace);
            Release(application);
            if (!retainOutput) TryDelete(path);
        }
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static void Release(object? value) {
        if (value == null || !Marshal.IsComObject(value)) return;
#pragma warning disable CA1416
        try { Marshal.FinalReleaseComObject(value); }
#pragma warning restore CA1416
        catch (InvalidComObjectException) { }
    }
}
