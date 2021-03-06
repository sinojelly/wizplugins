function OnKMButtonClicked() 
{
	var objApp = WizExplorerApp; 
	var objWindow = objApp.Window;
    var pluginPath = objApp.GetPluginPathByScriptFileName("SaveWiki.js");
    var helperFileName = pluginPath + "menu.htm";

    var rect = objWindow.GetToolButtonRect("document", "KMButton");
    var arr = rect.split(',');
    objWindow.ShowSelectorWindow(helperFileName, arr[0], arr[3], 120, 74, "");
}

function InitKMButton() 
{
	var pluginPath = objApp.GetPluginPathByScriptFileName("SaveWiki.js");
	var languangeFileName = pluginPath + "plugin.ini";
	var buttonText = objApp.LoadStringFromFile(languangeFileName,"strKM");
	objWindow.AddToolButton("document", "KMButton", buttonText, "", "OnKMButtonClicked");
}
InitKMButton();


