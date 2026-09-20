export function summarize(rows) {
  return rows.reduce((total, row) => total + row.value, 0);
}
