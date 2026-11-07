using System.Security.Cryptography;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Migration contract workloads; the shared PowerForge runner owns all measurements.</summary>
public abstract class MigrationWorkload {
    protected MigrationWorkload(byte[] input, int units) {
        Input = input;
        Units = units;
        InputSha256 = Convert.ToHexString(SHA256.HashData(input));
    }

    protected byte[] Input { get; }
    protected int Units { get; }
    protected byte[]? Output { get; set; }
    public string InputSha256 { get; }
    public long InputBytes => Input.LongLength;
    public virtual long OutputBytes => Output?.LongLength ?? 0;
    public int VerifiedUnits { get; protected set; }
    public string DiagnosticDigest { get; protected set; } = "";

    /// <summary>Creates one deterministic input outside the measured operation.</summary>
    public static MigrationWorkload Create(string profile, int units) => profile switch {
        "Word69" => new Word69Workload(units),
        "Biff4" => new BiffWorkload(units),
        "IWork09" => new IWork09Workload(units),
        "Dbx" => new DbxWorkload(units),
        "OfficeXml" => new OfficeXmlWorkload(units),
        _ => throw new ArgumentException("Unknown migration profile.", nameof(profile))
    };

    /// <summary>Runs the requested load/inspection or load/conversion/save public path once.</summary>
    public abstract void Execute(string operation);

    /// <summary>Verifies the complete ordered semantic result and saved artifact outside timing.</summary>
    public abstract void Validate();

    /// <summary>Releases result state after validation while retaining the immutable source.</summary>
    public virtual void ReleaseResults() => Output = null;
}
