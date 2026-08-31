namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffCodePageEncoding {
        private const int DefaultWindowsCodePage = 1252;
        private const ushort MacCodePageMarker = 32768;
        private const ushort WindowsCodePageMarker = 32769;

        static BiffCodePageEncoding() {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        internal static Encoding Resolve(BiffCodePageState declarations) {
            if (declarations is null) throw new ArgumentNullException(nameof(declarations));
            return Resolve(declarations.GetCodePageOrThrow());
        }

        internal static Encoding Resolve(ushort? codePage) {
            int effectiveCodePage = codePage switch {
                null or 0 => DefaultWindowsCodePage,
                MacCodePageMarker => 10000,
                WindowsCodePageMarker => DefaultWindowsCodePage,
                _ => codePage.Value
            };

            try {
                return Encoding.GetEncoding(
                    effectiveCodePage,
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
            } catch (Exception ex) when (ex is ArgumentException or NotSupportedException) {
                throw new InvalidDataException(
                    $"BIFF5 workbook code page {effectiveCodePage} is not supported.",
                    ex);
            }
        }
    }

    internal sealed class BiffCodePageState {
        private int _declarationCount;
        private ushort? _codePage;

        internal bool HasMalformedDeclaration { get; private set; }

        internal bool HasConflictingDeclarations { get; private set; }

        internal int? ProblemRecordOffset { get; private set; }

        internal bool HasDuplicateDeclarations => _declarationCount > 1;

        internal bool IsInvalid => HasMalformedDeclaration || HasConflictingDeclarations;

        internal string? InvalidReason => HasMalformedDeclaration
            ? "The BIFF5 workbook contains a malformed CodePage declaration, so byte strings cannot be decoded safely."
            : HasConflictingDeclarations
                ? "The BIFF5 workbook contains conflicting CodePage declarations, so byte strings cannot be decoded safely."
                : null;

        internal void ObserveMalformed(int recordOffset) {
            _declarationCount++;
            HasMalformedDeclaration = true;
            ProblemRecordOffset ??= recordOffset;
        }

        internal void Observe(ushort codePage, int recordOffset) {
            _declarationCount++;
            if (_codePage.HasValue && _codePage.Value != codePage) {
                HasConflictingDeclarations = true;
                ProblemRecordOffset ??= recordOffset;
                return;
            }

            _codePage ??= codePage;
        }

        internal ushort? GetCodePageOrThrow() {
            if (InvalidReason is string invalidReason) throw new InvalidDataException(invalidReason);

            return _codePage;
        }
    }
}
