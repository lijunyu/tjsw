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
	response.write "�ϴ��ɹ��� ����:"&file.FileName&"����С:"&file.FileSize&" [<a href=# onclick='redo()'>�����ϴ�</a>]"
	response.end
elseif file.fileSize<=0 then
	response.write "��ѡ���ļ��� [<a href=# onclick='redo()'>�����ϴ�</a>]"
	response.end
end if

set file=nothing
set upload=nothing
%>
</body> 
</html>