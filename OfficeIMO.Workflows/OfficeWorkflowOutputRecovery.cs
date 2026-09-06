namespace OfficeIMO.Workflows;

/// <summary>A validated local artifact retained when a provider's final contents need checking.</summary>
public sealed class OfficeWorkflowOutputRecovery {
    internal OfficeWorkflowOutputRecovery(string id, string name, string destination, string filePath,
        long length, string sha256, DateTimeOffset createdUtc) {
        Id = id; Name = name; Destination = destination; FilePath = filePath;
        Length = length; Sha256 = sha256; CreatedUtc = createdUtc;
    }
    /// <summary>Gets the store-local recovery identifier.</summary>
    public string Id { get; }
    /// <summary>Gets the intended provider filename.</summary>
    public string Name { get; }
    /// <summary>Gets the original provider location. Its current contents are not confirmed by this record.</summary>
    public string Destination { get; }
    /// <summary>Gets the local recovery copy. Verify it through the store before opening it.</summary>
    public string FilePath { get; }
    /// <summary>Gets the validated artifact length.</summary>
    public long Length { get; }
    /// <summary>Gets the validated artifact SHA-256.</summary>
    public string Sha256 { get; }
    /// <summary>Gets the time the recovery copy was created.</summary>
    public DateTimeOffset CreatedUtc { get; }
}
