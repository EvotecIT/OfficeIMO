/* Observe native stage precision without modifying native arithmetic or clipping. */
int office_probe_range_valid=1;
static int32_t office_range_check_value(int32_t value,int8_t bit) {
  int64_t limit=INT64_C(1)<<(bit-1);
  if(value< -limit || value>=limit) office_probe_range_valid=0;
  return value;
}
static void office_range_check_buf(int32_t stage,const int32_t *input,const int32_t *values,int32_t count,int8_t bit) {
  (void)stage;(void)input;
  for(int32_t i=0;i<count;i++) (void)office_range_check_value(values[i],bit);
}
#define range_check_value office_range_check_value
#define av1_range_check_buf office_range_check_buf
