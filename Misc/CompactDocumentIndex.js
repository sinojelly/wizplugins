var objApp = WizExplorerApp;
//
var objDatabase = objApp.Database;
//
try {
    objDatabase.SQLExecute("VACUUM", "");
    objApp.Window.ShowMessage(objApp.LoadPluginString("{C99F12F6-E038-4fe2-BC50-B4C217BAA97A}", "CompactSuccess"), "Wiz", 0x00000040);
}
catch (e) {
    objApp.Window.ShowMessage(e.message, "Wiz", 0x00000010);
}
