namespace OfficeIMO.Rtf;

internal sealed class RtfReadLimitGuard {
    private readonly RtfReadOptions _options;
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

    public RtfReadLimitGuard(RtfReadOptions options, CancellationToken cancellationToken) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _cancellationToken = cancellationToken;
    }

    public void CheckInputCharacters(long actual) => Check("RtfInputCharacterLimitExceeded", nameof(RtfReadOptions.MaxInputCharacters), actual, _options.MaxInputCharacters, -1);

    public void CheckInputBytes(long actual) => Check("RtfInputByteLimitExceeded", nameof(RtfReadOptions.MaxInputBytes), actual, _options.MaxInputBytes, -1);

    public void AddToken(int position) {
        Check("RtfTokenLimitExceeded", nameof(RtfReadOptions.MaxTokenCount), ++_tokenCount, _options.MaxTokenCount, position);
        CheckCancellation();
    }

    public void AddGroup(int position) => Check("RtfGroupLimitExceeded", nameof(RtfReadOptions.MaxGroupCount), ++_groupCount, _options.MaxGroupCount, position);

    public void AddTextCharacters(int count, int position) {
        if (count <= 0) return;
        _textCharacters += count;
        Check("RtfTextCharacterLimitExceeded", nameof(RtfReadOptions.MaxTextCharacters), _textCharacters, _options.MaxTextCharacters, position);
    }

    public void AddBinaryPayload(int count, int position) {
        Check("RtfBinaryPayloadLimitExceeded", nameof(RtfReadOptions.MaxBinaryBytesPerPayload), count, _options.MaxBinaryBytesPerPayload, position);
        _binaryBytes += count;
        Check("RtfTotalBinaryLimitExceeded", nameof(RtfReadOptions.MaxTotalBinaryBytes), _binaryBytes, _options.MaxTotalBinaryBytes, position);
    }

    public void BeginImage(int position) => Check("RtfImageCountLimitExceeded", nameof(RtfReadOptions.MaxImageCount), ++_imageCount, _options.MaxImageCount, position);

    public void AddImageBytes(ref long imageBytes, int count, int position) {
        if (count <= 0) return;
        imageBytes += count;
        Check("RtfImagePayloadLimitExceeded", nameof(RtfReadOptions.MaxImageBytesPerImage), imageBytes, _options.MaxImageBytesPerImage, position);
        _imageBytes += count;
        Check("RtfTotalImageLimitExceeded", nameof(RtfReadOptions.MaxTotalImageBytes), _imageBytes, _options.MaxTotalImageBytes, position);
    }

    public void BeginObject(int position) => Check("RtfObjectCountLimitExceeded", nameof(RtfReadOptions.MaxObjectCount), ++_objectCount, _options.MaxObjectCount, position);

    public void AddObjectBytes(ref long objectBytes, int count, int position) {
        if (count <= 0) return;
        objectBytes += count;
        Check("RtfObjectPayloadLimitExceeded", nameof(RtfReadOptions.MaxObjectBytesPerObject), objectBytes, _options.MaxObjectBytesPerObject, position);
        _objectBytes += count;
        Check("RtfTotalObjectLimitExceeded", nameof(RtfReadOptions.MaxTotalObjectBytes), _objectBytes, _options.MaxTotalObjectBytes, position);
    }

    public void AddSemanticBlock(int position) => Check("RtfSemanticBlockLimitExceeded", nameof(RtfReadOptions.MaxSemanticBlockCount), ++_semanticBlockCount, _options.MaxSemanticBlockCount, position);

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
