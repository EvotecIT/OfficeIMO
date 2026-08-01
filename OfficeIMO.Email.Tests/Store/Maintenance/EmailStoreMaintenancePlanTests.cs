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
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static byte[] Hash(byte[] value) { using SHA256 hash = SHA256.Create(); return hash.ComputeHash(value); }
}
