function OnSaveHistoryButtonClicked() 
{
	var objApp = WizExplorerApp; 
	var objWindow = objApp.Window;
    var pluginPath = objApp.GetPluginPathByScriptFileName("SaveHistory.js");
    var helperFileName = pluginPath + "commit.htm";

	objWindow.ShowHtmlDialog("commit", helperFileName, 350, 180, "", null); // add the last parameter 'null' for wiz 2.0
}

function InitSaveHistoryButton() 
{
	var pluginPath = objApp.GetPluginPathByScriptFileName("SaveHistory.js");
	var languangeFileName = pluginPath + "plugin.ini";
	var buttonText = objApp.LoadStringFromFile(languangeFileName,"strKM");
	objWindow.AddToolButton("document", "SaveHistoryButton", buttonText, "", "OnSaveHistoryButtonClicked");
}
InitSaveHistoryButton();


