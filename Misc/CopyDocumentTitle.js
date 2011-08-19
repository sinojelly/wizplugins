var objApp = WizExplorerApp;
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
//
var objDatabase = objApp.Database;
//
var selected_documents = objApp.Window.DocumentsCtrl.SelectedDocuments;
if (null == selected_documents
        || selected_documents.Count == 0) {
}
else {
    //
    var str = "";
    //
    for (var i = 0; i < selected_documents.Count; i++) {
        var doc = selected_documents.Item(i);
        str += doc.Title;
        str += "\r\n";
    }
    //
    objCommon.CopyTextToClipboard(str);
}
