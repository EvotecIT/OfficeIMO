using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Email.Store.Tests;

public sealed class EmailStoreMaintenancePlanTests {
    [Fact]
    public void MaintenancePlanningIsSourcePreservingAndFingerprintBound() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-maintenance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string message = Path.Combine(root, "message.eml");
        File.WriteAllText(message, "From: a@example.test\r\nTo: b@example.test\r\nSubject: safe\r\n\r\nbody", Encoding.ASCII);
        byte[] before = Hash(File.ReadAllBytes(message));
        try {
            using EmailStoreSession session = EmailStoreSession.Open(root);
            EmailStoreMaintenancePlan plan = session.PlanMaintenance();

            Assert.True(plan.PreservesSource);
            Assert.Equal(64, plan.SourceFingerprint.Length);
            Assert.Equal(EmailStoreMaintenanceAction.None, Assert.Single(plan.Recommendations).Action);
            Assert.Equal(before, Hash(File.ReadAllBytes(message)));
            Assert.Throws<NotSupportedException>(() =>
                ((IList<EmailStoreDiagnostic>)plan.Validation.Diagnostics).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<EmailStoreItemReference>)plan.Recovery.RecoveredItems).Clear());
            Assert.Throws<NotSupportedException>(() =>
                ((IList<EmailStoreMaintenanceRecommendation>)plan.Recommendations).Clear());
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void MaintenancePlanningRejectsSameLengthSourceMutationDuringScans() {
        byte[] message = Encoding.ASCII.GetBytes(
            "From sender@example.test Sat Jan 01 00:00:00 2022\n"
            + "From: a@example.test\nTo: b@example.test\nSubject: changing\n\nbody\n");
        using var source = new MutatingAfterFullReadStream(message);
        using EmailStoreSession session = EmailStoreSession.Open(source, "mailbox.mbox");
        source.Arm();

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
            () => session.PlanMaintenance());

        Assert.Contains("source changed", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] Hash(byte[] value) { using SHA256 hash = SHA256.Create(); return hash.ComputeHash(value); }

    private sealed class MutatingAfterFullReadStream : MemoryStream {
        private bool _armed;
        private bool _mutated;

        internal MutatingAfterFullReadStream(byte[] bytes)
            : base(bytes, 0, bytes.Length, writable: true, publiclyVisible: true) { }

        internal void Arm() => _armed = true;

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            if (_armed && !_mutated && read > 0 && Position == Length) {
                byte[] content = GetBuffer();
                content[checked((int)Length - 1)] = content[checked((int)Length - 1)] == (byte)'y'
                    ? (byte)'z'
                    : (byte)'y';
                _mutated = true;
            }
            return read;
        }
    }
}
