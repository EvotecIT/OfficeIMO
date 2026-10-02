/* Observation only: cropped post-superresolution pixels at both supported depths. */
#include "OfficeReconstructionProbe.inc"
static void office_trace_upscaled_frame(const AV1_COMMON *cm) {
  const YV12_BUFFER_CONFIG *b=&cm->cur_frame->buf;const int count=av1_num_planes(cm);
  const uint8_t *planes[3]={b->y_buffer,b->u_buffer,b->v_buffer};
  const int strides[3]={b->y_stride,b->uv_stride,b->uv_stride};
  const int high=(b->flags & YV12_FLAG_HIGHBITDEPTH)!=0;
  fprintf(stderr,"{\"width\":%d,\"height\":%d,\"planes\":[",cm->superres_upscaled_width,cm->height);
  for(int p=0;p<count;p++) {
    int sub=p?1:0,w=(cm->superres_upscaled_width+sub)>>sub,h=(cm->height+sub)>>sub;
    fprintf(stderr,"%s\"",p?",":"");
    for(int y=0;y<h;y++)for(int x=0;x<w;x++) {
      if(high) {uint16_t v=CONVERT_TO_SHORTPTR(planes[p])[y*strides[p]+x];fprintf(stderr,"%02x%02x",v&255,v>>8);}
      else fprintf(stderr,"%02x",planes[p][y*strides[p]+x]);
    }
    fprintf(stderr,"\"");
  }
  fprintf(stderr,"]}\n");
}
