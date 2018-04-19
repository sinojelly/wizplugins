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

function InsertPicture(pluginPath, work_dir, html, text, url_base)
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
	});

    var text_matchs = text.match(/<img_location name="(.*)?">/ig);

    if (html_matchs == null || text_matchs == null) // 没有图片
    {
        return text;
    }

	var image_dst_path = CombinePath(work_dir, url_base);
	CopyPictures(html_matchs, image_dst_path);
	
    var length = Math.min(html_matchs.length, text_matchs.length);

    for (i = 0; i < length; i++)
    {
        text = InsertLine( text, text_matchs[i] + "\n"
            , "![" + GetPictureName(text_matchs[i]) + "](" + GetPictureUrl(url_base, GetSrcFileName(html_matchs[i])) + ")\n");
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

function PrepareCopyPictures(local_path)
{
    var fso = objApp.CreateActiveXObject("Scripting.FileSystemObject"); // 如果用new ActiveXObject("XXXX"); 则会弹出对话框让用户确认
    CreateFolders(fso, local_path);
	return fso;
}

function CopyFile(fso, src_file, local_path)
{
	fso.CopyFile(src_file, CombinePath(local_path, GetSrcFileName(src_file)), true);
}

//https://stackoverflow.com/questions/423376/how-to-get-the-file-name-from-a-full-path-using-javascript
function GetSrcFileName(src_file) 
{
	var splitTest = function (str) {
        return str.split('\\').pop().split('/').pop();
	}
	return splitTest(src_file);
}

function CopyPictures(src_files, local_path)
{
	var i = 0;
    var fso = PrepareCopyPictures(local_path);
	for (i = 0; i < src_files.length; i++)
	{
		CopyFile(fso, src_files[i], local_path);
	}
}
