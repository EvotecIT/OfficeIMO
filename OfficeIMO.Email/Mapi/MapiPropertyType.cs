namespace OfficeIMO.Email;

/// <summary>MAPI property types used by MSG and TNEF artifacts.</summary>
public enum MapiPropertyType : ushort {
    /// <summary>Unspecified property type.</summary>
    Unspecified = 0x0000,
    /// <summary>Null value.</summary>
    Null = 0x0001,
    /// <summary>Signed 16-bit integer.</summary>
    Integer16 = 0x0002,
    /// <summary>Signed 32-bit integer.</summary>
    Integer32 = 0x0003,
    /// <summary>32-bit floating-point number.</summary>
    Floating32 = 0x0004,
    /// <summary>64-bit floating-point number.</summary>
    Floating64 = 0x0005,
    /// <summary>Signed 64-bit scaled currency value.</summary>
    Currency = 0x0006,
    /// <summary>Floating-point OLE Automation time.</summary>
    FloatingTime = 0x0007,
    /// <summary>32-bit error code.</summary>
    ErrorCode = 0x000A,
    /// <summary>Boolean value.</summary>
    Boolean = 0x000B,
    /// <summary>Embedded object or storage.</summary>
    Object = 0x000D,
    /// <summary>Signed 64-bit integer.</summary>
    Integer64 = 0x0014,
    /// <summary>Single-byte string.</summary>
    String8 = 0x001E,
    /// <summary>UTF-16LE string.</summary>
    Unicode = 0x001F,
    /// <summary>UTC FILETIME value.</summary>
    Time = 0x0040,
    /// <summary>GUID value.</summary>
    Guid = 0x0048,
    /// <summary>Binary value.</summary>
    Binary = 0x0102,
    /// <summary>Multiple signed 16-bit integers.</summary>
    MultipleInteger16 = 0x1002,
    /// <summary>Multiple signed 32-bit integers.</summary>
    MultipleInteger32 = 0x1003,
    /// <summary>Multiple 32-bit floating-point values.</summary>
    MultipleFloating32 = 0x1004,
    /// <summary>Multiple 64-bit floating-point values.</summary>
    MultipleFloating64 = 0x1005,
    /// <summary>Multiple currency values.</summary>
    MultipleCurrency = 0x1006,
    /// <summary>Multiple floating-point time values.</summary>
    MultipleFloatingTime = 0x1007,
    /// <summary>Multiple signed 64-bit integers.</summary>
    MultipleInteger64 = 0x1014,
    /// <summary>Multiple single-byte strings.</summary>
    MultipleString8 = 0x101E,
    /// <summary>Multiple UTF-16LE strings.</summary>
    MultipleUnicode = 0x101F,
    /// <summary>Multiple UTC FILETIME values.</summary>
    MultipleTime = 0x1040,
    /// <summary>Multiple GUID values.</summary>
    MultipleGuid = 0x1048,
    /// <summary>Multiple binary values.</summary>
    MultipleBinary = 0x1102
}
