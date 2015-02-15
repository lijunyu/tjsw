<html>
<body>
<!--#include FILE="upload_5xsoft.inc"--> 
<script>
function redo() {
	parent.picframe.location="./upload.asp";
}
function setpath(fname) {
	parent.document.getElementById('picpath').value='uploads/product/'+fname;
}
</script>
<%
set upload=new upload_5xsoft
set file=upload.file("file1")

if file.fileSize>0 then
	file.saveAs Server.mappath("../uploads/product/"&file.FileName)
	response.write "<script>setpath('"&file.FileName&"')</script>"
	response.write "上传成功！ 名称:"&file.FileName&"，大小:"&file.FileSize&" [<a href=# onclick='redo()'>重新上传</a>]"
	response.end
elseif file.fileSize<=0 then
	response.write "请选择文件！ [<a href=# onclick='redo()'>重新上传</a>]"
	response.end
end if

set file=nothing
set upload=nothing
%>
</body> 
</html>