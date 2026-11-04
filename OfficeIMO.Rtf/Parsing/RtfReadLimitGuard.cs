namespace OfficeIMO.Rtf;

internal sealed class RtfReadLimitGuard {
    private readonly RtfLoadOptions _options;
    private readonly CancellationToken _cancellationToken;
    private long _tokenCount;
    private long _groupCount;
    private long _textCharacters;
    private long _binaryBytes;
    private long _imageCount;
    private long _imageBytes;
    private long _objectCount;
    private long _objectBytes;
    private long _semanticBlockCount;
    private int _operationCount;

    public RtfReadLimitGuard(RtfLoadOptions options, CancellationToken cancellationToken) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (options.Limits == null) throw new ArgumentNullException(nameof(options), "RTF read limits are required.");
        _cancellationToken = cancellationToken;
    }
    internal bool DeferBinaryPayloads => _options.DeferBinaryPayloads;

    public void CheckInputCharacters(long actual) => Check("RtfInputCharacterLimitExceeded", nameof(RtfReadLimits.MaxInputCharacters), actual, _options.Limits.MaxInputCharacters, -1);

    public void CheckInputBytes(long actual) => Check("RtfInputByteLimitExceeded", nameof(RtfReadLimits.MaxInputBytes), actual, _options.Limits.MaxInputBytes, -1);

    public void AddToken(int position) {
        Check("RtfTokenLimitExceeded", nameof(RtfReadLimits.MaxTokens), ++_tokenCount, _options.Limits.MaxTokens, position);
        CheckCancellation();
    }

    public void AddGroup(int position) => Check("RtfGroupLimitExceeded", nameof(RtfReadLimits.MaxGroups), ++_groupCount, _options.Limits.MaxGroups, position);

    public void AddTextCharacters(int count, int position) {
        if (count <= 0) return;
        _textCharacters += count;
        Check("RtfTextCharacterLimitExceeded", nameof(RtfReadLimits.MaxTextCharacters), _textCharacters, _options.Limits.MaxTextCharacters, position);
    }

    public void AddBinaryPayload(int count, int position) {
        Check("RtfBinaryPayloadLimitExceeded", nameof(RtfReadLimits.MaxBinaryPayloadBytes), count, _options.Limits.MaxBinaryPayloadBytes, position);
        _binaryBytes += count;
        Check("RtfTotalBinaryLimitExceeded", nameof(RtfReadLimits.MaxTotalBinaryBytes), _binaryBytes, _options.Limits.MaxTotalBinaryBytes, position);
    }

    public void BeginImage(int position) => Check("RtfImageCountLimitExceeded", nameof(RtfReadLimits.MaxImages), ++_imageCount, _options.Limits.MaxImages, position);

    public void AddImageBytes(ref long imageBytes, int count, int position) {
        if (count <= 0) return;
        imageBytes += count;
        Check("RtfImagePayloadLimitExceeded", nameof(RtfReadLimits.MaxImageBytes), imageBytes, _options.Limits.MaxImageBytes, position);
        _imageBytes += count;
        Check("RtfTotalImageLimitExceeded", nameof(RtfReadLimits.MaxTotalImageBytes), _imageBytes, _options.Limits.MaxTotalImageBytes, position);
    }

    public void BeginObject(int position) => Check("RtfObjectCountLimitExceeded", nameof(RtfReadLimits.MaxObjects), ++_objectCount, _options.Limits.MaxObjects, position);

    public void AddObjectBytes(ref long objectBytes, int count, int position) {
        if (count <= 0) return;
        objectBytes += count;
        Check("RtfObjectPayloadLimitExceeded", nameof(RtfReadLimits.MaxObjectBytes), objectBytes, _options.Limits.MaxObjectBytes, position);
        _objectBytes += count;
        Check("RtfTotalObjectLimitExceeded", nameof(RtfReadLimits.MaxTotalObjectBytes), _objectBytes, _options.Limits.MaxTotalObjectBytes, position);
    }

    public void AddSemanticBlock(int position) => Check("RtfSemanticBlockLimitExceeded", nameof(RtfReadLimits.MaxSemanticBlocks), ++_semanticBlockCount, _options.Limits.MaxSemanticBlocks, position);

    public void CheckCancellation() {
        if ((_operationCount++ & 0x3FF) == 0) {
            _cancellationToken.ThrowIfCancellationRequested();
        }
    }

    public void ThrowIfCancellationRequested() => _cancellationToken.ThrowIfCancellationRequested();

    private static void Check(string code, string source, long actual, long? limit, int position) {
        if (limit.HasValue && actual > limit.Value) {
            throw new RtfReadLimitException(code, $"RTF input exceeded {source} ({actual} > {limit.Value}).", source, actual, limit.Value, position);
        }
    }
}
