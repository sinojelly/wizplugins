// 1. Find "<img_location name="picture_name">"
// 2. Find "<img src="index_files/94b373d7-5131-4385-b055-ac4921a54465.png">" in html content, picture file name is not fixed.
// 3. Copy the picture to WizDocs/images/article_name/
// 4. Insert ![picture_name](path/to/picture) after img_location.

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

function GetPictureUrl(url_base, html_tag) // html_tag is filename
{
    return CombineUrl(url_base, html_tag);
}

function GetPictureName(html_tag)
{
    return html_tag.substr(20, html_tag.length - 22);  // 去掉<img_location name="  和 ">
}

//https://stackoverflow.com/questions/423376/how-to-get-the-file-name-from-a-full-path-using-javascript
function GetSrcFileName(src_file) 
{
    var splitTest = function (str) {
        return str.split('\\').pop().split('/').pop();
    }
    return splitTest(src_file);
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

function InsertPicture(html_matchs, text_matchs, text)
{
    var length = Math.min(html_matchs.length, text_matchs.length);

    for (i = 0; i < length; i++)
    {
        text = InsertLine( text, text_matchs[i] + "\n"
            , "![" + GetPictureName(text_matchs[i]) + "](" + GetPictureUrl(pic_path, GetSrcFileName(html_matchs[i])) + ")\n");
    }
    return text;
}

function saveContent(comment, text) 
{
    try {
        objComm.SaveTextToFile(file_path, text, "UTF-8");
    }catch (e) {
        alert("Write file (" + file_path + ") failed!");
        return;
    }
    
    var image_dst_path = CombinePath(work_dir, pic_path)
    CopyPictures(objDoc, image_dst_path);
        
    objComm.RunExe(pluginPath + commit_cmd, "\""+ tool_path + "\" \""+work_dir + "\" \"" + comment + "\"", true);  // return 0 - sucsess
}

function ConstructMarkdownContentAndSave(comment, html, text)
{
    //var html_matchs = html.match(/<IMG src="index_files\/(.*)?.(png|jpg|jpeg|bmp|gif|ico)">/ig);
    var html_matchs = [];
    var scriptFile = pluginPath + "get_imgs.js";
    var objBrowser = objApp.Window.CurrentDocumentBrowserObject;
    objBrowser.ExecuteScriptFile(scriptFile, function(ret){
        if(ret && ret.length > 0){
            for(var i = 0; i < ret.length; i++){
                html_matchs.push(ret[i]);  
            }
        }
        
        var text_matchs = text.match(/<img_location name="(.*)?">/ig);
        if (html_matchs.length > 0 && text_matchs.length > 0) // 有图片
        {
            text = InsertPicture(html_matchs, text_matchs, text);
        }
        saveContent(comment, text);
        
        objApp.Window.CloseHtmlDialog(WizChromeBrowser, "ok");
    });
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
}
