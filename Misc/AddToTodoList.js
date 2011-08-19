
function formatInt(val) {
    if (val < 10)
        return "0" + val;
    else
        return "" + val;
}

function DateToStr(dt) {
    return "" + dt.getFullYear() + "-" + formatInt(dt.getMonth() + 1) + "-" + formatInt(dt.getDate());
}

function GetTodayDateTimeString() {
    var dtNow = new Date();
    return DateToStr(dtNow) + " 00:00:00";
}

var objApp = WizExplorerApp;
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
//
var objDatabase = objApp.Database;
//
var docTodo = objDatabase.GetTodoDocument(GetTodayDateTimeString());
if (docTodo != null)
{
    var selected_documents = objApp.Window.DocumentsCtrl.SelectedDocuments;
    if (null == selected_documents
            || selected_documents.Count == 0) {
    }
    else {
        //
        var items = docTodo.TodoItems;
        //
        for (var i = 0; i < selected_documents.Count; i++) {
            var doc = selected_documents.Item(i);
            //
            var title = doc.Title;
            //
            var todo = objApp.CreateWizObject("WizKMCore.WizTodoItem");
            todo.Text = title;
            todo.Prior = 0;
            todo.Complete = 0;
            todo.LinkedDocumentGUIDs = [doc.GUID];
            todo.DateCreated = doc.DateCreated;
            todo.DateModified = doc.DateModified;
            //
            items.Add(todo);
        }
        //
        docTodo.TodoItems = items;
    }
}
