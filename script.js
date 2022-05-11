function onEdit(e) {
  var col = e.range.getColumn();
  if (col === 1) {
    var row = e.range.getRow();
    e.source
      .getActiveSheet()
      .getRange(row, 2)
      .setValue(e.user + ' ' + new Date());
  }
}
