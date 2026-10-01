/* Observation only: called after all single-threaded tiles, before any frame filters. */
#include <stdio.h>
static int office_leaves,office_skip,office_copy,office_fractional_copy;
static void office_trace_reconstruction_leaf(const MB_MODE_INFO *m) {
  office_leaves++;office_skip+=m->skip_txfm!=0;office_copy+=m->use_intrabc!=0;
  if(m->use_intrabc && ((m->mv[0].as_mv.row&8) || (m->mv[0].as_mv.col&8)))office_fractional_copy++;
}
static void office_trace_reconstructed_frame(const AV1_COMMON *cm) {
  const YV12_BUFFER_CONFIG *b = &cm->cur_frame->buf;
  const uint8_t *planes[3] = {b->y_buffer,b->u_buffer,b->v_buffer};
  const int strides[3] = {b->y_stride,b->uv_stride,b->uv_stride};
  const int count = av1_num_planes(cm);
  fprintf(stderr,"{\"width\":%d,\"height\":%d,\"level\":%d,\"miCols\":%d,\"miRows\":%d,\"tileCols\":%d,\"tileRows\":%d,\"leaves\":%d,\"skipLeaves\":%d,\"copyLeaves\":%d,\"fractionalCopyLeaves\":%d,\"planes\":[",
          cm->width,cm->height,cm->seq_params->seq_level_idx[0],cm->mi_params.mi_cols,cm->mi_params.mi_rows,cm->tiles.cols,cm->tiles.rows,office_leaves,office_skip,office_copy,office_fractional_copy);
  for(int p=0;p<count;p++) {
    int s=p?1:0,w=cm->mi_params.mi_cols*4>>s,h=cm->mi_params.mi_rows*4>>s;
    fprintf(stderr,"%s\"",p?",":"");
    for(int y=0;y<h;y++) for(int x=0;x<w;x++) fprintf(stderr,"%02x",planes[p][y*strides[p]+x]);
    fprintf(stderr,"\"");
  }
  fprintf(stderr,"]}\n");
}
