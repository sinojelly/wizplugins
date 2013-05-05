/**
 * @project wiz sheet
 * @author sinojelly@gmail.com
 * $Id: wizsheet.js 2013-05-05 16:14:000 sinojelly $
 * Licensed under MIT
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */
function updateWizSheet() {
    var htmltext = $('table#jS_0_0').get(0).outerHTML;
    htmltext = '<div id="sheetParent">' + htmltext + '</div>';

    var html = $('html').get(0).cloneNode(true);  // Clone HTML using DOM API
    var $html = $(html);                     // Make jQuery object from cloned HTML

    $('#sheetParent', $html).replaceWith(htmltext);
    $html = removeUseless($html);

    // TODO: should add title and script to $html head. from current document's $('head')
    showDebugInfo($html.get(0).outerHTML);
    /*alert('暂时还未实现完成!');var objDoc = window.external.Window.CurrentDocumentHtmlDocument;objDoc.UpdateDocument3($html.get(0).outerHTML, 0x0002);*/
}

function checkVersion() {
    var browser=navigator.appName
    var b_version=navigator.appVersion
    var version=parseFloat(b_version)
    alert("Browser ："+ browser + "  Version ："+ version);
}

function dumpObj(obj) {
    var description = "";
    for (var i in obj) {
        var property = obj[i];
        description += i + " = " + property + "\n";

    }
    return (description);
}

function showDebugInfo(infoString) {      // for debuging
    var htmltext = infoString;
    if (htmltext == null) {
        //htmltext =document.all[0].innerHTML;
        //htmltext = $("table#jS_0_0").get(0).outerHTML;
        //htmltext = $('html')[0].outerHTML;
        htmltext = $("#sheetParent").get(0).outerHTML;   //include sheetParent div node
    }

    var info = $('div#debugInfo')[0];
    if (info) {  // exist, remove first
        info.remove(0);
    }
    info = $('<div id=\"debugInfo\"></div>')[0];
    $('body').append(info);
    info.innerText = htmltext; // 直接用 .text(htmltext) 赋值会丢失格式
}

/*
 jQuery1830009167320928771594="72"
 sizset="true" sizcache005113121418754335="2 7 0"
 sizset="false" sizcache005113121418754335="11 22 22"
 */
function removeJQueryAttri(input) {
    return input.replace(/((jQuery|jquery)\d*=\"[\s\S]*?\")|(sizset=\"(true|false)\")|(sizcache\d*=\"[\s\S]*?\")/g, "");
}
function removeUselessAttri(obj) {
    var inhtml = $('body', obj)[0].innerHTML;
    inhtml = removeJQueryAttri(inhtml);
    $('body', obj)[0].innerHTML = inhtml;
    return obj;
}

// obj is jquery object
function removeUseless(obj) {
    $('.jSMenu', obj).remove();
    $('div#sheetParent ~ div', obj).remove();
    return removeUselessAttri(obj);
}
