<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Index.aspx.cs" Inherits="WebApplication1.Index" %>

<!DOCTYPE html>

<html>
<head>
    <title>利用JQuery的ajax请求实现文件上传</title>
    <script src="Scripts/jquery-3.3.1.min.js"></script>
</head>
<body>
    <div>
        <input type="file" name="file1" id="file1" />
        <button type="button" id="submitId">点击上传</button>
    </div>
    <div>
        <input type="button" id="CURDWord" value="开始操作word" />
    </div>
</body>
</html>
<script>
    $("#submitId").click(function () {
        var formData = new FormData();
        formData.append("myfile", document.getElementById("file1").files[0]);   
        $.ajax({
            url: "/FileUpLoad.ashx?action=Upfileload",
            type: "POST",
            data: formData,
            /**
            *必须false才会自动加上正确的Content-Type
            */
            contentType: false,
            /**
            * 必须false才会避开jQuery对 formdata 的默认处理
            * XMLHttpRequest会对 formdata 进行正确的处理
            */
            processData: false,
            success: function (data) {
                var res = JSON.parse(data);
                if (res.state == 'success') {
                    alert("上传成功！");
                } else {
                     alert("上传失败！");
                }
            },
            error: function () {
                alert("上传失败！");
                $("#imgWait").hide();
            }
        });
    });

    $("#CURDWord").click(function () {
        $.ajax({
            url: "/FileUpLoad.ashx?action=CURDWord",
            contentType:"application/json;charset=utf-8",
            dataType: "json",
            data: "",
            success: function (data) {
                if (data == 200) {
                    alert("操作完成");
                } else {
                    alert("出错了");
                }
                
            }
        });
    });
</script>
