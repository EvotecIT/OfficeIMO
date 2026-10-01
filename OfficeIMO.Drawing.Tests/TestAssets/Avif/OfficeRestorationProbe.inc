/* Observation only: true cropped output and completed restoration-unit metadata. */
#include "OfficeReconstructionProbe.inc"
static void office_trace_restored_frame(const AV1_COMMON *cm) {
  const YV12_BUFFER_CONFIG *b=&cm->cur_frame->buf;const int count=av1_num_planes(cm);
  const uint8_t *planes[3]={b->y_buffer,b->u_buffer,b->v_buffer};
  const int strides[3]={b->y_stride,b->uv_stride,b->uv_stride};
  fprintf(stderr,"{\"width\":%d,\"height\":%d,\"planes\":[",cm->superres_upscaled_width,cm->height);
  for(int p=0;p<count;p++) {
    int sub=p?1:0,w=(cm->superres_upscaled_width+sub)>>sub,h=(cm->height+sub)>>sub;
    fprintf(stderr,"%s\"",p?",":"");
    for(int y=0;y<h;y++)for(int x=0;x<w;x++)fprintf(stderr,"%02x",planes[p][y*strides[p]+x]);
    fprintf(stderr,"\"");
  }
  fprintf(stderr,"],\"restoration\":[");int first=1;
  for(int p=0;p<count;p++) {
    const RestorationInfo *r=&cm->rst_info[p];
    if(r->frame_restoration_type==RESTORE_NONE)continue;
    for(int i=0;i<r->num_rest_units;i++) {
      const RestorationUnitInfo *u=&r->unit_info[i];
      fprintf(stderr,"%s{\"plane\":%d,\"row\":%d,\"col\":%d,\"size\":%d,\"type\":%d,\"set\":%d,\"x0\":%d,\"x1\":%d}",
        first?"":",",p,i/r->horz_units,i%r->horz_units,r->restoration_unit_size,u->restoration_type,
        u->restoration_type==RESTORE_SGRPROJ?u->sgrproj_info.ep:0,u->restoration_type==RESTORE_SGRPROJ?u->sgrproj_info.xqd[0]:0,
        u->restoration_type==RESTORE_SGRPROJ?u->sgrproj_info.xqd[1]:0);first=0;
    }
  }
  fprintf(stderr,"]}\n");
}
