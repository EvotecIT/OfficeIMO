using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1IntraPredictor {
    /// <summary>Predicts CfL from a bounded reconstructed luma region, extending its final row/column at frame edges.</summary>
    internal OfficeAv1Prediction PredictChromaFromLuma(int size,int alpha,ushort dc,ushort[] luma,int stride,
        int sourceWidth,int sourceHeight,int subX,int subY) {
        var d=Dimensions(size);int w=d.Width,h=d.Height;
        if(luma==null) throw new ArgumentNullException(nameof(luma));
        if(w>32 || h>32 || (uint)subX>1 || (uint)subY>1 || subY>subX || alpha< -16 || alpha>16 ||
           sourceWidth<(1<<subX) || sourceWidth>64 || sourceHeight<(1<<subY) || sourceHeight>64 ||
           (sourceWidth&subX)!=0 || (sourceHeight&subY)!=0 || stride<sourceWidth || stride>64 ||
           luma.Length>4096 || (long)(sourceHeight-1)*stride+sourceWidth>luma.Length)
            throw new FormatException("Invalid AV1 chroma-from-luma source or scaling.");
        ValidateSample(dc);
        int sum=0;
        for(int y=0;y<h;y++) {
            _cancellation.ThrowIfCancellationRequested();int sy=Math.Min(y<<subY,sourceHeight-(1<<subY));
            for(int x=0;x<w;x++) {
                int sx=Math.Min(x<<subX,sourceWidth-(1<<subX)),value=0;
                for(int dy=0;dy<=subY;dy++) for(int dx=0;dx<=subX;dx++) {int sample=luma[(sy+dy)*stride+sx+dx];ValidateSample(sample);value+=sample;}
                value<<=3-subX-subY;_luma[y*w+x]=value;sum+=value;
            }
        }
        int average=(sum+w*h/2)/(w*h);var output=new ushort[w*h];
        for(int y=0;y<h;y++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=0;x<w;x++) {int scaled=alpha*(_luma[y*w+x]-average);scaled=scaled<0?-((-scaled+32)>>6):(scaled+32)>>6;output[y*w+x]=Clip(dc+scaled);}
        }
        _cancellation.ThrowIfCancellationRequested();return new OfficeAv1Prediction(w,h,output);
    }

    /// <summary>Maps one transform's leaf-relative palette indices to immutable prediction samples.</summary>
    internal OfficeAv1Prediction PredictPalette(int size,int plane,OfficeAv1Palette palette,int offsetX,int offsetY) {
        var d=Dimensions(size);int w=d.Width,h=d.Height;
        if(palette==null) throw new ArgumentNullException(nameof(palette));
        if((uint)plane>=3) throw new FormatException("Invalid AV1 palette plane.");
        int width=plane==0?palette.Width:palette.ChromaWidth,height=plane==0?palette.Height:palette.ChromaHeight;
        int colors=plane==0?palette.SizeY:palette.SizeUv;
        if(colors<2 || colors>8 || offsetX<0 || offsetY<0 || offsetX>width-w || offsetY>height-h)
            throw new FormatException("Invalid AV1 palette prediction region.");
        for(int i=0;i<colors;i++) ValidateSample(palette.Color(plane,i));
        var output=new ushort[w*h];
        for(int y=0;y<h;y++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int x=0;x<w;x++) {
                int index=palette.Index(plane>0,y+offsetY,x+offsetX);
                if(index>=colors) throw new FormatException("Invalid AV1 palette prediction index.");
                output[y*w+x]=palette.Color(plane,index);
            }
        }
        _cancellation.ThrowIfCancellationRequested();return new OfficeAv1Prediction(w,h,output);
    }
}
