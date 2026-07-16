using OfficeIMO.Email;
using System.Runtime.InteropServices;

namespace OfficeIMO.Email.Store.Tests;

public sealed class OutlookPstWriterInteropTests {
    [OutlookInteropFact]
    public void Generated_unicode_pst_can_be_mounted_read_and_removed_by_classic_outlook() {
#pragma warning disable CA1416
        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
#pragma warning restore CA1416
        Assert.NotNull(outlookType);

        string? retainedPath = Environment.GetEnvironmentVariable(
            "OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP_OUTPUT");
        string path = string.IsNullOrWhiteSpace(retainedPath)
            ? Path.Combine(Path.GetTempPath(),
                string.Concat("officeimo-outlook-interop-", Guid.NewGuid().ToString("N"), ".pst"))
            : Path.GetFullPath(retainedPath!);
        object? application = null;
        object? nameSpace = null;
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
            if (!string.IsNullOrWhiteSpace(retainedPath)) {
                File.Copy(path, string.Concat(path, ".before-outlook"), overwrite: true);
            }
            application = Activator.CreateInstance(outlookType);
            Assert.NotNull(application);
            dynamic outlook = application!;
            nameSpace = outlook.GetNamespace("MAPI");
            dynamic mapi = nameSpace!;
            mapi.AddStoreEx(path, 2); // OlStoreType.olStoreUnicode
            dynamic stores = mapi.Stores;
            for (int index = 1; index <= stores.Count; index++) {
                dynamic candidate = stores.Item(index);
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
            root = null;
        } finally {
            Release(root);
            Release(store);
            Release(nameSpace);
            Release(application);
            if (string.IsNullOrWhiteSpace(retainedPath)) {
                try { if (File.Exists(path)) File.Delete(path); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    private static void Release(object? value) {
        if (value == null || !Marshal.IsComObject(value)) return;
#pragma warning disable CA1416
        try { Marshal.FinalReleaseComObject(value); }
#pragma warning restore CA1416
        catch (InvalidComObjectException) { }
    }
}
