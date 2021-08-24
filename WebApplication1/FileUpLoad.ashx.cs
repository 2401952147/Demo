using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace WebApplication1
{
    /// <summary>
    /// FileUpLoad 的摘要说明
    /// </summary>
    public class FileUpLoad : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string action = context.Request["action"].ToString();
            var _str = "";
            switch (action)
            {
                case "Upfileload":
                    Upfileload(context);//文件上传
                    break;
                case "CURDWord":
                    _str = CURDWord(context);//文件上传
                    break;
                default:
                    break;
            }

            context.Response.Write(_str);
        }

        public void Upfileload(HttpContext context)
        {
            //Upload为自定义的文件夹，在项目中创建
            string UrlPath = "/Public/File/";
            //此为限制文件格式
            string[] ExtentsfileName = new string[] { ".doc", ".xls", ".png", ".jpg" };
            //保存文件名
            string name = "";
            if (context.Request.Files.Count > 0)
            {
                foreach (string fn in context.Request.Files)
                {
                    var file = context.Request.Files[fn];
                    name = file.FileName.ToString();
                    var extenfilename = Path.GetExtension(file.FileName);
                    //判断 路径是否存在
                    string path = HttpContext.Current.Server.MapPath(UrlPath);
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    if (ExtentsfileName.Contains(extenfilename.ToLower()))
                    {
                        string urlfile = UrlPath + name;
                        string filepath = HttpContext.Current.Server.MapPath(urlfile);
                        file.SaveAs(filepath);
                    }

                    //格式不正确
                    else
                    {
                        context.Response.Write("格式不正确");
                    }
                }

                //上传成功
                context.Response.Write("{\"state\":\"success\",\"msg\":\"成功\"}");
                //return 200;
            }

            //上传失败
            else
            {
                context.Response.Write("{\"state\":\"fail\",\"msg\":\"失败\"}");
                //return 100;
            }
        }

        //word书签操作
        public string CURDWord(HttpContext context)
        {
            MSWord.Application wordApp;               //Word应用程序变量 
            MSWord.Document wordDoc;
            killWinWordProcess();
            wordApp = new ApplicationClass();
            wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Visible = false;
            object missing = System.Reflection.Missing.Value;
            //HttpContext.Current.Server.MapPath(Path.Combine("/upload/download/", "MyWord_Print.pdf"));
            object templateName = /*wordApp.StartupPath + */HttpContext.Current.Server.MapPath(Path.Combine("/Public/File/", "ReportModel_Stand2.doc"));//最终的word文档需要写入的位置
            object ModelName = /*wordApp.StartupPath + */HttpContext.Current.Server.MapPath(Path.Combine("/Public/File/", "ReportModel_Stand.doc")); ;//word模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;
            wordDoc = wordApp.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板

            //在书签处插入文字
            object oStart = "PatName";//word中的书签名 
            Range range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//表格插入位置 
            range.Text = "这里是说明内容aaaaaaaaa";//在书签处插入文字内容

            //在书签处插入表格
            oStart = "PatInfo";//word中的书签名 
            range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//表格插入位置      
            MSWord.Table tab_Pat = wordDoc.Tables.Add(range, 2, 4, ref missing, ref missing);//开辟一个2行4列的表格
            tab_Pat.Range.Font.Size = 10.5F;
            tab_Pat.Range.Font.Bold = 0;

            tab_Pat.Columns[1].Width = 50;
            tab_Pat.Columns[2].Width = 65;
            tab_Pat.Columns[3].Width = 40;
            tab_Pat.Columns[4].Width = 40;

            tab_Pat.Cell(1, 1).Range.Text = "病历号";
            tab_Pat.Cell(1, 2).Range.Text = "PatientNO";
            tab_Pat.Cell(1, 3).Range.Text = "身高";
            tab_Pat.Cell(1, 4).Range.Text = "Height";

            tab_Pat.Cell(2, 1).Range.Text = "姓名";
            tab_Pat.Cell(2, 2).Range.Text = "PatientName";
            tab_Pat.Cell(2, 3).Range.Text = "体重";
            tab_Pat.Cell(2, 4).Range.Text = "Weight";


            //保存word
            object format = WdSaveFormat.wdFormatDocument;//保存格式 
            wordDoc.SaveAs(ref templateName, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //关闭wordDoc，wordApp对象              
            object SaveChanges = WdSaveOptions.wdSaveChanges;
            object OriginalFormat = WdOriginalFormat.wdOriginalDocumentFormat;
            object RouteDocument = false;
            wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            wordApp.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);

            return "200";
        }

        public void killWinWordProcess()
        {
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            foreach (System.Diagnostics.Process process in processes)
            {
                bool b = process.MainWindowTitle == "";
                if (process.MainWindowTitle == "")
                {
                    process.Kill();
                }
            }
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}
