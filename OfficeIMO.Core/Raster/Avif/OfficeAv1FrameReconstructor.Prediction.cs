using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1FrameReconstructor {
    private OfficeAv1Prediction Predict(OfficeAv1TransformBlock b) {
        int p=b.Plane,sub=p==0?0:1;
        int x=b.X-(_block.Region.MiCol>>sub)*4,y=b.Y-(_block.Region.MiRow>>sub)*4;
        if((p==0?_block.Palette.SizeY:_block.Palette.SizeUv)>0)
            return _predictor.PredictPalette(b.Size,p,_block.Palette,x,y);
        var mode=p==0?_block.Modes.YMode:_block.Modes.UvMode;
        bool cfl=mode==OfficeAv1IntraMode.ChromaFromLuma;
        var prediction=_predictor.Predict(b.Size,cfl?OfficeAv1IntraMode.Dc:mode,
            p==0?_block.Modes.AngleDeltaY:_block.Modes.AngleDeltaUv,Edges(b),
            _sequence.IntraEdgeFilter,SmoothNeighbors(p),p==0?_block.Palette.FilterMode:-1);
        if(!cfl) return prediction;
        int sx=b.X*2,sy=b.Y*2,w=Math.Min(64,_maxLumaX-sx),h=Math.Min(64,_maxLumaY-sy);
        if(w<2 || h<2) throw new FormatException("AV1 CfL has no reconstructed luma extent.");
        for(int row=0;row<h;row++) {
            _cancellation.ThrowIfCancellationRequested();for(int col=0;col<w;col++) _luma[row*w+col]=_pixels[0][(sy+row)*_stride+sx+col];
        }
        return _predictor.PredictChromaFromLuma(b.Size,p==1?_block.Modes.CflAlphaU:_block.Modes.CflAlphaV,
            prediction.Value(0),_luma,w,w,h,1,1);
    }

    private int CopySample(OfficeAv1TransformBlock b,int x,int y) {
        int sub=b.Plane==0?0:1,stride=_stride>>sub;
        int sx=b.X+x+(_block.Motion.Col>>(3+sub)),sy=b.Y+y+(_block.Motion.Row>>(3+sub));
        int dx=sub==1 && (_block.Motion.Col&8)!=0?1:0,dy=sub==1 && (_block.Motion.Row&8)!=0?1:0;
        // At half chroma positions use the normative 2-tap filter, rounding once after both dimensions.
        if(sx<(_tile.MiColStart*4>>sub) || sy<(_tile.MiRowStart*4>>sub) ||
           sx+dx>=(_tile.MiColEnd*4>>sub) || sy+dy>=(_tile.MiRowEnd*4>>sub))
            throw new FormatException("AV1 copy interpolation footprint exceeds its tile.");
        var pixels=_pixels[b.Plane];int value=pixels[sy*stride+sx];
        if(dx!=0) value+=pixels[sy*stride+sx+1];
        if(dy!=0) {value+=pixels[(sy+1)*stride+sx];if(dx!=0)value+=pixels[(sy+1)*stride+sx+1];}
        int bits=dx+dy;return bits==0?value:(value+(1<<(bits-1)))>>bits;
    }
}
