using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace OfficeIMO.Drawing.Wic.Interop;

// Slot order and native signatures follow the Windows SDK wincodec.h.
// Opaque COM pointers stay explicit so every reference is released before COM uninitialization.
[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("ec5ec8a9-c395-4314-9c77-54d7a935ff70")]
internal partial interface IWicImagingFactory {
    [PreserveSig] int CreateDecoderFromFilename(nint name, nint vendor, uint access, uint cache, out nint decoder);
    [PreserveSig] int CreateDecoderFromStream(nint stream, nint vendor, uint cache, out nint decoder);
    [PreserveSig] int CreateDecoderFromFileHandle(nuint handle, nint vendor, uint cache, out nint decoder);
    [PreserveSig] int CreateComponentInfo(in Guid clsid, out nint info);
    [PreserveSig] int CreateDecoder(in Guid format, nint vendor, out nint decoder);
    [PreserveSig] int CreateEncoder(in Guid format, nint vendor, out nint encoder);
    [PreserveSig] int CreatePalette(out nint palette);
    [PreserveSig] int CreateFormatConverter(out nint converter);
    [PreserveSig] int CreateBitmapScaler(out nint scaler);
    [PreserveSig] int CreateBitmapClipper(out nint clipper);
    [PreserveSig] int CreateBitmapFlipRotator(out nint rotator);
    [PreserveSig] int CreateStream(out nint stream);
    [PreserveSig] int CreateColorContext(out nint context);
    [PreserveSig] int CreateColorTransformer(out nint transformer);
    [PreserveSig] int CreateBitmap(uint width, uint height, in Guid format, uint cache, out nint bitmap);
    [PreserveSig] int CreateBitmapFromSource(nint source, uint cache, out nint bitmap);
    [PreserveSig] int CreateBitmapFromSourceRect(nint source, uint x, uint y, uint width, uint height, out nint bitmap);
    [PreserveSig] int CreateBitmapFromMemory(uint width, uint height, in Guid format, uint stride, uint size, nint buffer, out nint bitmap);
    [PreserveSig] int CreateBitmapFromHBITMAP(nint bitmap, nint palette, uint alpha, out nint result);
    [PreserveSig] int CreateBitmapFromHICON(nint icon, out nint bitmap);
    [PreserveSig] int CreateComponentEnumerator(uint types, uint options, out nint enumerator);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("135ff860-22b7-4ddf-b0f6-218f4f299a43")]
internal partial interface IWicStream {
    // ISequentialStream and IStream slots precede IWICStream's own methods.
    [PreserveSig] int Read(nint buffer, uint count, nint read);
    [PreserveSig] int Write(nint buffer, uint count, nint written);
    [PreserveSig] int Seek(long move, uint origin, nint position);
    [PreserveSig] int SetSize(ulong size);
    [PreserveSig] int CopyTo(nint stream, ulong count, nint read, nint written);
    [PreserveSig] int Commit(uint flags);
    [PreserveSig] int Revert();
    [PreserveSig] int LockRegion(ulong offset, ulong count, uint type);
    [PreserveSig] int UnlockRegion(ulong offset, ulong count, uint type);
    [PreserveSig] int Stat(nint stat, uint flags);
    [PreserveSig] int Clone(out nint stream);
    [PreserveSig] int InitializeFromIStream(nint stream);
    [PreserveSig] int InitializeFromFilename(nint filename, uint access);
    [PreserveSig] int InitializeFromMemory(nint buffer, uint size);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("9edde9e7-8dee-47ea-99df-e6faf2ed44bf")]
internal partial interface IWicBitmapDecoder {
    [PreserveSig] int QueryCapability(nint stream, out uint capability);
    [PreserveSig] int Initialize(nint stream, uint cache);
    [PreserveSig] int GetContainerFormat(out Guid format);
    [PreserveSig] int GetDecoderInfo(out nint info);
    [PreserveSig] int CopyPalette(nint palette);
    [PreserveSig] int GetMetadataQueryReader(out nint reader);
    [PreserveSig] int GetPreview(out nint preview);
    [PreserveSig] int GetColorContexts(uint count, nint contexts, out uint actual);
    [PreserveSig] int GetThumbnail(out nint thumbnail);
    [PreserveSig] int GetFrameCount(out uint count);
    [PreserveSig] int GetFrame(uint index, out nint frame);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("00000120-a8f2-4877-ba0a-fd2b6645fb94")]
internal partial interface IWicBitmapSource {
    [PreserveSig] int GetSize(out uint width, out uint height);
    [PreserveSig] int GetPixelFormat(out Guid format);
    [PreserveSig] int GetResolution(out double x, out double y);
    [PreserveSig] int CopyPalette(nint palette);
    [PreserveSig] int CopyPixels(nint rectangle, uint stride, uint size, nint pixels);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("00000301-a8f2-4877-ba0a-fd2b6645fb94")]
internal partial interface IWicFormatConverter : IWicBitmapSource {
    [PreserveSig] int Initialize(nint source, in Guid format, uint dither, nint palette, double alphaThreshold, uint paletteType);
    [PreserveSig] int CanConvert(in Guid sourceFormat, in Guid destinationFormat, out int canConvert);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("00000100-0000-0000-c000-000000000046")]
internal partial interface IEnumUnknown {
    [PreserveSig] int Next(uint count, out nint item, out uint fetched);
}

[GeneratedComInterface(Options = ComInterfaceOptions.ComObjectWrapper)]
[Guid("23bc3f0a-698b-4357-886b-f24d50671334")]
internal partial interface IWicComponentInfo {
    [PreserveSig] int GetComponentType(out uint type);
    [PreserveSig] int GetCLSID(out Guid clsid);
}
