function onEdit(e) {
  var col = e.range.getColumn();
  if (col === 1) {
    var row = e.range.getRow();
    e.source
      .getActiveSheet()
      .getRange(row, 2)
      .setValue(getCurrentUserEmail() + ' ' + new Date());
  }
}

function getCurrentUserEmail() {
  var userEmail = Session.getActiveUser().getEmail();
  if (userEmail === '' || !userEmail || userEmail === undefined) {
    userEmail = PropertiesService.getUserProperties().getProperty('userEmail');
    if (!userEmail) {
      var protection = SpreadsheetApp.getActive().getRange('A1').protect();
      protection.removeEditors(protection.getEditors());
      var editors = protection.getEditors();
      if (editors.length === 2) {
        var owner = SpreadsheetApp.getActive().getOwner();
        editors.splice(editors.indexOf(owner), 1);
      }
      userEmail = editors[0];
      protection.remove();
      PropertiesService.getUserProperties().setProperty('userEmail', userEmail);
    }
  }
  return userEmail;
}
