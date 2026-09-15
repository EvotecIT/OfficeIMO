using OfficeIMO.Email;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace OfficeIMO.Email.Store.Tests;

public sealed class OutlookPstWriterInteropTests {
    [OutlookInteropFact]
    public void Generated_unicode_pst_passes_scanpst_without_repair_or_byte_changes() {
        string? scanPst = FindScanPst();
        Assert.False(string.IsNullOrWhiteSpace(scanPst),
            "Microsoft Inbox Repair Tool (SCANPST.EXE) is not installed.");
        string path = Path.Combine(Path.GetTempPath(),
            string.Concat("officeimo-scanpst-interop-", Guid.NewGuid().ToString("N"), ".pst"));
        string logPath = Path.ChangeExtension(path, ".log");
        try {
            WriteInteropStore(path, empty: false);
            byte[] before = ComputeHash(path);
            var start = new System.Diagnostics.ProcessStartInfo {
                FileName = scanPst!,
                Arguments = string.Concat("-file \"", path,
                    "\" -force -silent -no repair -log replace"),
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (System.Diagnostics.Process? process =
                System.Diagnostics.Process.Start(start)) {
                Assert.NotNull(process);
                Assert.True(process!.WaitForExit(60_000),
                    "SCANPST did not finish within one minute.");
            }

            Assert.Equal(before, ComputeHash(path));
            Assert.True(File.Exists(logPath), "SCANPST did not produce its validation log.");
            string log = File.ReadAllText(logPath, Encoding.Unicode);
            Assert.DoesNotContain("!!", log, StringComparison.Ordinal);
            Assert.DoesNotContain("Start Repairing", log, StringComparison.OrdinalIgnoreCase);
        } finally {
            TryDelete(logPath);
            TryDelete(path);
        }
    }

    [OutlookInteropFact]
#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
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
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        bool completed = thread.Join(TimeSpan.FromMinutes(2));
        if (!completed && string.IsNullOrWhiteSpace(retainedPath)) TryDelete(path);
        Assert.True(completed,
            "Classic Outlook interoperability did not finish within two minutes.");
        if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
    }

#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    private static void RunInterop(string path, bool retainOutput) {
        Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
        Assert.NotNull(outlookType);

        object? application = null;
        object? nameSpace = null;
        object? stores = null;
        object? store = null;
        object? root = null;
        try {
            WriteInteropStore(path, string.Equals(Environment.GetEnvironmentVariable(
                "OFFICEIMO_EMAIL_STORE_OUTLOOK_INTEROP_EMPTY"), "1", StringComparison.Ordinal));
            if (retainOutput) {
                File.Copy(path, string.Concat(path, ".before-outlook"), overwrite: true);
            }
            application = Activator.CreateInstance(outlookType);
            Assert.NotNull(application);
            dynamic outlook = application!;
            nameSpace = outlook.GetNamespace("MAPI");
            dynamic mapi = nameSpace!;
            stores = mapi.Stores;
            dynamic outlookStores = stores!;
            int originalStoreCount = outlookStores.Count;
            mapi.AddStoreEx(path, 2); // OlStoreType.olStoreUnicode
            int mountedStoreCount = outlookStores.Count;
            Assert.Equal(originalStoreCount + 1, mountedStoreCount);
            // Stores added through AddStoreEx are appended. Restrict lookup to
            // the newly added range so an unrelated unavailable profile store
            // cannot make this interoperability check fail.
            for (int index = originalStoreCount + 1; index <= mountedStoreCount; index++) {
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
                Assert.Contains("OfficeIMO classic Outlook semantic body",
                    Convert.ToString(item.Body), StringComparison.Ordinal);
                Assert.Equal(1, (int)item.Recipients.Count);
                dynamic recipient = item.Recipients.Item(1);
                Assert.Equal("recipient@example.test",
                    Convert.ToString(recipient.Address));
                Assert.Equal(1, (int)item.Attachments.Count);
                dynamic attachment = item.Attachments.Item(1);
                Assert.Equal("outlook-evidence.txt",
                    Convert.ToString(attachment.FileName));
                Release(attachment);
                Release(recipient);
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

    private static void WriteInteropStore(string path, bool empty) {
        using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
            new EmailStorePstWriterOptions("OfficeIMO Interop"))) {
            if (!empty) {
                string folder = writer.AddFolder("OfficeIMO Synthetic");
                var document = new EmailDocument {
                    Subject = "OfficeIMO synthetic interoperability item",
                    MessageClass = "IPM.Note",
                    From = new EmailAddress("sender@example.test", "OfficeIMO sender")
                };
                document.Body.Text = "OfficeIMO classic Outlook semantic body";
                document.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
                    new EmailAddress("recipient@example.test", "OfficeIMO recipient")));
                byte[] attachment = Encoding.UTF8.GetBytes(
                    "OfficeIMO classic Outlook attachment evidence");
                document.Attachments.Add(new EmailAttachment {
                    FileName = "outlook-evidence.txt",
                    ContentType = "text/plain",
                    Content = attachment,
                    Length = attachment.LongLength
                });
                writer.AddItem(folder, document);
            }
            writer.Complete();
        }
    }

    private static string? FindScanPst() {
        string[] roots = {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
        };
        return roots.Where(root => !string.IsNullOrWhiteSpace(root))
            .Select(root => Path.Combine(root, "Microsoft Office", "root", "Office16", "SCANPST.EXE"))
            .FirstOrDefault(File.Exists);
    }

    private static byte[] ComputeHash(string path) {
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (SHA256 sha256 = SHA256.Create()) return sha256.ComputeHash(stream);
    }

#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    private static void Release(object? value) {
        if (value == null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); }
        catch (InvalidComObjectException) { }
    }
}
