using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;

string assetPath = Path.Combine(Path.GetTempPath(), "officeimo-c2pa-aot-" + Guid.NewGuid().ToString("N") + ".png");
string missingToolPath = Path.Combine(Path.GetTempPath(), "officeimo-c2pa-tool-" + Guid.NewGuid().ToString("N"), "c2patool");
try {
    File.WriteAllBytes(assetPath, new byte[] { 0x89, 0x50, 0x4E, 0x47 });
    var verifier = new C2paToolProvenanceVerifier(missingToolPath);
    OfficeProvenanceVerificationResult result = verifier.Verify(assetPath);
    if (result.Status != OfficeProvenanceVerificationStatus.ProviderUnavailable) {
        throw new InvalidOperationException($"Expected ProviderUnavailable for an absent c2patool executable, found {result.Status}.");
    }
} finally {
    File.Delete(assetPath);
}

Console.WriteLine("PASS | optional OfficeIMO.Provenance.C2pa provider-unavailable contract passed from NativeAOT.");
