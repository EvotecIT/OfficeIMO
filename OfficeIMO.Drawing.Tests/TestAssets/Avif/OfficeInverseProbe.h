/* Opt-in native observation only. The original AOM transforms and pixel clipping are unchanged. */
static int32_t *office_probe_residual;
static uint16_t *office_probe_origin;
static uint16_t office_probe_add(uint16_t *destination,int value,int bd) {
  if(office_probe_residual) office_probe_residual[destination-office_probe_origin]=value;
  return highbd_clip_pixel_add(*destination,value,bd);
}
