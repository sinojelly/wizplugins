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
