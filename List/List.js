
function OnListButtonClicked() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("List.js");
    var bookmarksListHtmlFileName = pluginPath + "List.htm";
    //
    var rect = objWindow.GetToolButtonRect("document", "ListButton");
    var arr = rect.split(',');
    objWindow.ShowSelectorWindow(bookmarksListHtmlFileName, arr[0], arr[3], 300, 500, "");
}
function InitListButton() {
    var pluginPath = objApp.GetPluginPathByScriptFileName("List.js");
    var languangeFileName = pluginPath + "plugin.ini";
    var buttonText = objApp.LoadStringFromFile(languangeFileName, "strList");
    objWindow.AddToolButton("document", "ListButton", buttonText, "", "OnListButtonClicked");
}

InitListButton();
