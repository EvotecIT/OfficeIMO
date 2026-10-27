using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.Versioning;

namespace OfficeIMO.Drawing.Wic.Interop;

[SupportedOSPlatform("windows")]
internal static partial class WicNative {
    internal static readonly Guid FactoryClass = new("cacaf262-9370-4615-a13b-9f5539da4c0a");
    internal static readonly Guid FactoryInterface = new("ec5ec8a9-c395-4314-9c77-54d7a935ff70");
    internal static readonly Guid HeifDecoderClass = new("e9a4a80a-44fe-4de4-8971-7150b10a5199");
    internal static readonly Guid DecoderInterface = new("9edde9e7-8dee-47ea-99df-e6faf2ed44bf");
    internal static readonly Guid Rgba32 = new("f5c7ad2d-6a8d-43dd-a7a8-a29935261ae9");
    internal static readonly StrategyBasedComWrappers Wrappers = new();

    [LibraryImport("ole32.dll")]
    internal static partial int CoInitializeEx(nint reserved, uint apartment);

    [LibraryImport("ole32.dll")]
    internal static partial void CoUninitialize();

    [LibraryImport("ole32.dll")]
    internal static partial int CoCreateInstance(in Guid clsid, nint outer, uint context, in Guid iid, out nint instance);

    internal static void Check(int hr) {
        if (hr < 0) throw new COMException($"WIC operation failed (HRESULT 0x{hr:X8}).", hr);
    }

    internal static WicObject<IWicImagingFactory> CreateFactory() {
        Check(CoCreateInstance(FactoryClass, 0, 1, FactoryInterface, out nint pointer));
        return new WicObject<IWicImagingFactory>(pointer);
    }
}

/// <summary>Balances this call's COM initialization without changing an existing STA apartment.</summary>
[SupportedOSPlatform("windows")]
internal readonly struct WicApartment : IDisposable {
    private readonly bool _uninitialize;

    public WicApartment() {
        int hr = WicNative.CoInitializeEx(0, 0); // COINIT_MULTITHREADED
        if (hr != unchecked((int)0x80010106)) WicNative.Check(hr); // RPC_E_CHANGED_MODE: retain caller's STA
        _uninitialize = hr >= 0; // S_FALSE also increments the COM initialization count.
    }

    public void Dispose() {
        if (_uninitialize) WicNative.CoUninitialize();
    }
}

/// <summary>Owns one native output reference and an uncached generated wrapper on the current thread.</summary>
[SupportedOSPlatform("windows")]
internal sealed class WicObject<T> : IDisposable where T : class {
    private ComObject? _wrapper;
    internal nint Pointer { get; private set; }
    internal T Value { get; }

    internal WicObject(nint pointer) {
        if (pointer == 0) throw new COMException("WIC returned a null COM pointer.", unchecked((int)0x80004003));
        Pointer = pointer;
        try {
            _wrapper = (ComObject)WicNative.Wrappers.GetOrCreateObjectForComInstance(pointer, CreateObjectFlags.UniqueInstance);
            Value = (T)(object)_wrapper;
        } catch {
            Dispose();
            throw;
        }
    }

    public void Dispose() {
        _wrapper?.FinalRelease();
        _wrapper = null;
        if (Pointer != 0) {
            Marshal.Release(Pointer);
            Pointer = 0;
        }
    }
}
