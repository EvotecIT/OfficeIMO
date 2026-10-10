using System.Globalization;

namespace OfficeIMO.Chm;

public sealed partial class ChmDocument {
    private void ReadMetadata(CancellationToken token) {
        ChmEntry? system = FindEntry("/#SYSTEM");
        var fields = new List<(int Code, byte[] Data)>();
        if (system != null) {
            byte[] data = system.GetBytes();
            ChmBinary.Range(data, 0, 4);
            uint version = ChmBinary.U32(data, 0);
            if (version != 2 && version != 3) throw ChmBinary.Error("METADATA_UNSUPPORTED", "Only #SYSTEM versions 2 and 3 are supported.");
            for (int position = 4; position < data.Length;) {
                token.ThrowIfCancellationRequested();
                ChmBinary.Range(data, position, 4);
                int code = ChmBinary.U16(data, position), length = ChmBinary.U16(data, position + 2);
                position += 4; ChmBinary.Range(data, position, length);
                if (code == 4 && length >= 4) LocaleId = ChmBinary.U32(data, position);
                // Retain only the handful of named metadata values; unknown records remain available in the raw entry.
                if (code == 0 || code == 1 || code == 2 || code == 3 || code == 9) {
                    var bytes = new byte[length]; Buffer.BlockCopy(data, position, bytes, 0, length); fields.Add((code, bytes));
                }
                position += length;
            }
        }
        int codePage = 1252;
        try { codePage = CultureInfo.GetCultureInfo(checked((int)LocaleId)).TextInfo.ANSICodePage; }
        catch (ArgumentException) { Diagnostic("CHM_LOCALE_UNKNOWN", "The help-book locale is unknown; undeclared text uses Windows-1252."); }
        catch (OverflowException) { Diagnostic("CHM_LOCALE_UNKNOWN", "The help-book locale is unknown; undeclared text uses Windows-1252."); }
        string? htmlLabel = codePage switch { 932 => "shift_jis", 936 => "gbk", 949 => "euc-kr", 950 => "big5", 874 => "windows-874", 65001 => "utf-8", _ => null };
        Encoding? resolved = htmlLabel == null ? null : _options.EncodingProvider.ResolveLabel(htmlLabel);
        resolved = resolved ?? _options.EncodingProvider.ResolveLabel("windows-" + codePage.ToString(CultureInfo.InvariantCulture))
            ?? _options.EncodingProvider.ResolveLabel(codePage.ToString(CultureInfo.InvariantCulture));
        if (_options.TextEncoding == null && resolved == null) Diagnostic("CHM_ENCODING_FALLBACK", "The charset provider cannot resolve the help-book code page; undeclared text uses Windows-1252.");
        _encoding = _options.TextEncoding ?? resolved ?? _options.EncodingProvider.ResolveLabel("windows-1252") ?? Encoding.UTF8;
        foreach (var field in fields) {
            int length = Array.IndexOf(field.Data, (byte)0); if (length < 0) length = field.Data.Length;
            string value = _encoding.GetString(field.Data, 0, length);
            switch (field.Code) {
                case 0: _contentsPath = value; break;
                case 1: _indexPath = value; break;
                case 2: DefaultTopic = value; break;
                case 3: Title = value; break;
                case 9: Compiler = value; break;
            }
        }
        if (DefaultTopic != null && FindEntry(DefaultTopic) == null)
            Diagnostic("CHM_DEFAULT_TOPIC_MISSING", "The declared default topic is absent or external.", DefaultTopic);
    }
}
