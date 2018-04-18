// 1. Find "<img_location name="picture_name">"
// 2. Find "<img src="index_files/94b373d7-5131-4385-b055-ac4921a54465.png">" in html content, picture file name is not fixed.
// 3. Copy the picture to WizDocs/images/article_name/
// 4. Insert ![picture_name](path/to/picture) after img_location.

// var objApp = window.external;
// var objComm = objApp.CreateWizObject("WizKMControls.WizCommonUI");

function CombinePath(path, file)
{
    path = path.replace(/\//g, "\\");
    file = file.replace(/\//g, "\\");

    if (path.charAt(path.length - 1) == '\\')
    {
        return path + file;
    }

    return path + "\\" + file;
}

function CombineUrl(url_base, file)
{
    if (url_base.charAt(url_base.length - 1) == '/')
    {
        return url_base + file;
    }

    return url_base + "/" + file;
}

function GetPictureUrl(url_base, html_tag)
{
    var file = html_tag.substr(22, html_tag.length - 24);  // 去掉<IMG src="index_files/  和 ">
    return CombineUrl(url_base, file);
}

function GetPictureName(html_tag)
{
    return html_tag.substr(20, html_tag.length - 22);  // 去掉<img_location name="  和 ">
}

// 把str中的src行后面添加一行dst
function InsertLine(str, src, dst)
{
    var index = str.indexOf(src, 0);
    if(index == -1) {
        return str;
    }
    var result = str.substr(0, index + src.length);
    result += dst;
    result += str.substr(index + src.length);
    return result;
}

function InsertPicture(html, text, url_base)
{
    var html_matchs = html.match(/<IMG src="index_files\/(.*)?.(png|jpg|jpeg|bmp|gif|ico)">/ig);

    var text_matchs = text.match(/<img_location name="(.*)?">/ig);

    if (html_matchs == null || text_matchs == null) // 没有图片
    {
        return text;
    }

    var length = Math.min(html_matchs.length, text_matchs.length);

    for (i = 0; i < length; i++)
    {
        text = InsertLine( text, text_matchs[i] + "\n"
            , "![" + GetPictureName(text_matchs[i]) + "](" + GetPictureUrl(url_base, html_matchs[i]) + ")\n");
    }

    return text;
}


// 把str中的src替换为dst，src后面紧跟except的除外
function ReplaceExceptOld(str, src, dst, except)
{
    var index = 0;
    var index2 = 0;

    while (index < str.length)
    {
        index2 = index;
        index = str.indexOf(src, index2);
        if (index != str.indexOf(src + except, index2))
        {
            var result = str.substr(0, index);
            result += dst;
            result += str.substr(index + src.length);
            return result;
        }
        index = index + src.length;
    }

    return str;
}

function InsertPictureOld(html, text, url_base)
{
    var html_matchs = html.match(/<IMG src="index_files\/(.*)?.(png|jpg|jpeg|bmp|gif|ico)">/ig);

    var text_re = /((参见下图)|(如图)|(如图所示)|(如下图所示))+[^\n]*\n/g;
    var text_matchs = text.match(text_re);

    if (html_matchs == null || text_matchs == null) // 没有图片
    {
        return text;
    }

    var length = Math.min(html_matchs.length, text_matchs.length);

    var auto_pic = "<wiki:comment>Picture added by Wiz2Wiki.</wiki:comment>";

    for (i = 0; i < length; i++)
    {
        text = ReplaceExcept( text, text_matchs[i]
            , text_matchs[i] + auto_pic + "\n  " + GetPictureUrl(url_base, html_matchs[i]) + "\n"
            , auto_pic + "\n  http://");
    }

    return text;
}

/////////////////////////////////////////////////////////////////////////////
////// The following functions should be called in Wiz enviroment.
/////////////////////////////////////////////////////////////////////////////

function CreateFolders(fso, path)
{
    var i = 0;
    var arr = [];
    while (!fso.FolderExists(path))
    {
        arr[i++] = path;
        path = path.substring(0, path.lastIndexOf("\\"));
    }

    for (i--; i >= 0; i--)
    {
        fso.CreateFolder(arr[i]);
    }
}


function CopyPictures(doc, local_path)
{
    var fso = objApp.CreateActiveXObject("Scripting.FileSystemObject"); // 如果用new ActiveXObject("XXXX"); 则会弹出对话框让用户确认

    var temp_file = objComm.GetATempFileName(".html");
    doc.SaveToHtml(temp_file, 1);

    var src_dir = temp_file.substr(0, temp_file.length - 5) + "_files";

    if (!fso.FolderExists(src_dir))
    {
        return;
    }

    CreateFolders(fso, local_path);

    fso.CopyFolder(src_dir, local_path, true);

    // 避免拷贝模板附带的文件，凡是带有[]的，都认为是模板的文件，不拷贝。
    //var currentFolder = fso.GetFolder(src_dir);
    /*var fileList = new Enumerator(currentFolder.files);  // only work in IE, not work in Chrome, because of security
    var aFile;
    for (; !fileList.atEnd(); fileList.moveNext())
    {
        aFile=fileList.item();
        if (/[^\[\]]+.(png|jpg|jpeg|bmp|gif|ico)/i.test(aFile.Name))
        {
            fso.CopyFile(aFile.Path, CombinePath(local_path, aFile.Name), true);
        }
    }*/

    //fso.CopyFolder(src_dir, CombinePath(local_path, "images/" + bare_file_name));
}

function Commit()
{
    var objWindow = objApp.Window;
    var objDoc = objWindow.CurrentDocument; //获得当前正在浏览的Wiz文档(WizDocument)

    var pluginPath = objApp.GetPluginPathByScriptFileName("SaveWiki.js");

    var file_name = objDoc.ParamValue("gocowiki_file_name");
    var local_path = objDoc.ParamValue("gocowiki_local_path");
    var remote_path = objDoc.ParamValue("gocowiki_remote_path");
    var commit_cmd = objDoc.ParamValue("gocowiki_commit_cmd");

    if ((file_name == "") || (local_path == "")
        || (remote_path == "") || (commit_cmd == ""))
    {
        alert("file_name local_path remote_path and commit_cmd can not be null, please set first!");
        var settings_dialog = pluginPath + "settings.htm";
        objWindow.ShowHtmlDialog("settings", settings_dialog, 450, 300, "", null);
        return;
    }

    var bare_file_name = file_name.substr(0, file_name.length - 5); // 去掉.wiki
    CopyPictures(objDoc, local_path, bare_file_name);

    var url_base = CombineUrl(remote_path, "images/" + bare_file_name);
    var text = InsertPicture(objDoc.GetHtml(), objDoc.GetText(0), url_base)
    text = text.replace(/\n\s*\n/g, "\n\n");  // 空行不能有空白字符, 否则它不会在wiki显示为换行，排版会比较乱

    //objComm.SaveTextToFile(CombinePath(local_path, file_name), text, "UTF-8");
    var file_path = CombinePath(local_path, file_name);
    try {
        objComm.SaveTextToFile(file_path, text, "UTF-8");
    }catch (e) {
        alert("Write file (" + file_path + ") failed!");
        return;
    }
    objComm.RunExe(pluginPath + commit_cmd /*"hg_commit.bat"*/, local_path + " \"" + textComment.value + "\"", true);  // return 0 - sucsess
}

function CloseDialog(ret)
{
    if (ret == 1)
    {
        Commit();
    }

    //var objApp = new ActiveXObject("WizExplorer.WizExplorerApp");
    var objApp = window.external;
    objApp.Window.CloseHtmlDialog(document, ret);
}
