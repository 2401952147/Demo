using System;
using System.IO;
using System.Linq;
using System.Web;

//操作word书签用
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Collections;
using System.Text;
using System.Net.Mail;
using MailMessage = System.Net.Mail.MailMessage;
using System.Net;
using Demo.Common;

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
                //case "CURDWord":
                //    _str = CURDWord(context);//word书签操作
                //    break;
                case "ExcelOperationClass":
                    _str = ExcelOperationClass(context);//导出到excel
                    break;
                case "QQ_email":
                    _str = QQ_email();//发送验证码
                    break;
                case "ExcelTo":
                    _str = ExcelTo(context);//Excel导入
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
            }

            //上传失败
            else
            {
                context.Response.Write("{\"state\":\"fail\",\"msg\":\"失败\"}");
            }
        }

        /// <summary>
        /// word书签操作
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 杀掉windows线程
        /// </summary>
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

        /// <summary>
        /// 导出到excel
        /// </summary>
        public string ExcelOperationClass(HttpContext context)
        {
            #region 引用
            //using NPOI.HSSF.UserModel;
            //using NPOI.HPSF;
            //using NPOI.POIFS.FileSystem;
            //using NPOI.SS.UserModel;
            //using NPOI.SS.Util;
            #endregion

            #region 创建Excel文件并写入表头
            //创建一个新的excel文件
            HSSFWorkbook book = new HSSFWorkbook();
            ISheet sheet = book.CreateSheet("sheet1");
            //创建一行 也就是在sheet1这个工作区创建一行 在NPOI中只有先创建才能后使用
            IRow row = sheet.CreateRow(0);//--索引从0开始
            for (int i = 0; i < 3; i++)
            {
                //设置单元格的宽度
                sheet.SetColumnWidth(i, 30 * 360);
            }
            sheet.SetColumnWidth(3, 30 * 156);
            sheet.SetColumnWidth(4, 30 * 156);
            sheet.SetColumnWidth(5, 30 * 156);
            sheet.SetColumnWidth(6, 30 * 156);

            //定义一个样式，迎来设置样式属性
            ICellStyle setborder = book.CreateCellStyle();

            //设置单元格上下左右边框线 但是不包括最外面的一层
            setborder.BorderLeft = BorderStyle.Thin;
            setborder.BorderRight = BorderStyle.Thin;
            setborder.BorderBottom = BorderStyle.Thin;
            setborder.BorderTop = BorderStyle.Thin;

            //文字水平和垂直对齐方式
            setborder.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            setborder.Alignment = HorizontalAlignment.Center;//水平居中
            setborder.WrapText = true;//自动换行

            //再定义一个样式，用来设置最上面标题行的样式
            ICellStyle setborderdeth = book.CreateCellStyle();

            //设置单元格上下左右边框线 但是不包括最外面的一层
            setborderdeth.BorderLeft = BorderStyle.Thin;
            setborderdeth.BorderRight = BorderStyle.Thin;
            setborderdeth.BorderBottom = BorderStyle.Thin;
            setborderdeth.BorderTop = BorderStyle.Thin;

            //定义一个字体样式
            IFont font = book.CreateFont();
            //将字体设为红色
            font.Color = IndexedColors.Red.Index;
            //font.FontHeightInPoints = 17;
            //将定义的font样式给到setborderdeth样式中
            setborderdeth.SetFont(font);

            //文字水平和垂直对齐方式
            setborderdeth.VerticalAlignment = VerticalAlignment.Center;//垂直居中
            setborderdeth.Alignment = HorizontalAlignment.Center;//水平居中
            setborderdeth.WrapText = true;  //自动换行

            //设置第一行单元格的高度为25
            row.HeightInPoints = 20;
            //设置单元格的值
            row.CreateCell(0).SetCellValue("流程");
            //将style属性给到这个单元格
            row.GetCell(0).CellStyle = setborderdeth;
            row.CreateCell(1).SetCellValue("二级目录");
            row.GetCell(1).CellStyle = setborderdeth;
            row.CreateCell(2).SetCellValue("任务");
            row.GetCell(2).CellStyle = setborderdeth;
            row.CreateCell(3).SetCellValue("得分");
            row.GetCell(3).CellStyle = setborderdeth;
            row.CreateCell(4).SetCellValue("个人分");
            row.GetCell(4).CellStyle = setborderdeth;
            row.CreateCell(5).SetCellValue("团队分");
            row.GetCell(5).CellStyle = setborderdeth;
            row.CreateCell(6).SetCellValue("总分");
            row.GetCell(6).CellStyle = setborderdeth;

            #endregion

            ////循环的导出到excel的每一行
            //for (int i = 0; i < Data.Count; i++)
            //{
            //    //每循环一次，就新增一行  索引从0开始 所以第一次循环CreateRow(1) 前面已经创建了标题行为0
            //    IRow row1 = sheet.CreateRow(i + 1);
            //    row1.HeightInPoints = 21;
            //    //给新加的这一行创建第一个单元格，并且给这第一个单元格设置值 以此类推...
            //    row1.CreateCell(0).SetCellValue(Convert.ToString(Data[i].Number));
            //    //先获取这一行的第一个单元格，再给其设置样式属性 以此类推...
            //    row1.GetCell(0).CellStyle = setborder;
            //    row1.CreateCell(1).SetCellValue(Data[i].ShopName);
            //    row1.GetCell(1).CellStyle = setborder;
            //    row1.CreateCell(2).SetCellValue(Convert.ToString(Data[i].Price));
            //    row1.GetCell(2).CellStyle = setborder;
            //    row1.CreateCell(3).SetCellValue(Data[i].ShopType);
            //    row1.GetCell(3).CellStyle = setborder;
            //    row1.CreateCell(4).SetCellValue(Convert.ToString(Data[i].Date));
            //    row1.GetCell(4).CellStyle = setborder;
            //}

            #region 合并
            int firstRow = 0;//起始行
            int lastRow = 0;//结束行
            int firstcol = 0;//起始列
            int lastcol = 0;//结束列
            sheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstcol, lastcol));
            #endregion

            #region 导出返回路径
            string path = "/upload/download/";
            if (Directory.Exists(path) == false)//如果不存在就创建_paths文件夹
            {
                Directory.CreateDirectory(path);
            }
            //string filename = "实训记录.xls";
            //using (FileStream sm = File.OpenWrite(HttpContext.Current.Server.MapPath(Path.Combine(path, filename))))
            //{
            //    _sb2.Append("{");
            //    _sb2.AppendFormat("\"code\":0");
            //    _sb2.AppendFormat(",\"msg\":\"\"");
            //    _sb2.Append(",\"data\": [");
            //    _sb2.Append("{");
            //    _sb2.AppendFormat("\"PathUrl\":\"{0}\"", path + filename);
            //    _sb2.Append("}");
            //    _sb2.Append("]");
            //    _sb2.Append("}");
            //    book.Write(sm);
            //    context.Response.Write(Message.Json("成功", _sb2.ToString()));
            //}
            #endregion

            return "成功";
        }

        public string QQ_email()
        {
            //实例化一个发送邮件类。
            MailMessage mailMessage = new MailMessage();
            //发件人邮箱地址，方法重载不同，可以根据需求自行选择。
            mailMessage.From = new MailAddress("1765079118@qq.com");
            //收件人邮箱地址。
            mailMessage.To.Add(new MailAddress("2401952147@qq.com"));
            //邮件标题。
            mailMessage.Subject = "验证码";
            //邮件内容。
            mailMessage.Body = "您好，您的验证码为：663245";

            //实例化一个SmtpClient类。
            SmtpClient client = new SmtpClient();
            //在这里我使用的是qq邮箱，所以是smtp.qq.com，如果你使用的是126邮箱，那么就是smtp.126.com。
            client.Host = "smtp.qq.com";
            //使用安全加密连接。
            client.EnableSsl = true;
            //不和请求一块发送。
            client.UseDefaultCredentials = false;
            //验证发件人身份(发件人的邮箱，邮箱里的生成授权码);
            client.Credentials = new NetworkCredential("1765079118@qq.com", "hpk54188");
            //发送
            client.Send(mailMessage);
            return "发送成功";
        }

        #region 数组去重去空 ArrayList的示例应用
        /// 方法名：DelArraySame
        /// 功能： 删除数组中重复的元素
        /// </summary>
        /// <param name="TempArray">所要检查删除的数组</param>
        /// <returns>返回数组</returns>
        public string[] DelArraySame(string[] TempArray)
        {
            ArrayList nStr = new ArrayList();
            for (int i = 0; i < TempArray.Length; i++)
            {
                if (!nStr.Contains(TempArray[i]) && !string.IsNullOrEmpty(TempArray[i]))
                {
                    nStr.Add(TempArray[i]);
                }
            }
            string[] newStr = (string[])nStr.ToArray(typeof(string));
            return newStr;
        }
        #endregion

        public string ExcelTo(HttpContext context)
        {
            string _json = "";
            string Extension = "";
            float contentLength = 0;
            string fullPath = "";

            HttpFileCollection hfc = context.Request.Files;
            if (hfc.Count > 0)
            {
                HttpPostedFile file = context.Request.Files[0];
                HttpPostedFile hpf = context.Request.Files[0];

                contentLength = hpf.ContentLength; //文件大小
                string contentType = hpf.ContentType;  //文件类型
                if (hpf.ContentLength > 0)
                {
                    //获取扩展名 
                    Extension = Path.GetExtension(hpf.FileName);

                    string ext = Path.GetExtension(context.Request.Files[0].FileName);
                    //string mapPath = HttpContext.Current.Server.MapPath(context.Request.ApplicationPath);
                    string dir = HttpContext.Current.Server.MapPath("/Public/File/" + DateTime.Now.ToString("yyyy-MM-dd"));
                    DirectoryInfo dirInfo = Directory.CreateDirectory(dir);
                    //fullPath = dir + "/" + Guid.NewGuid().ToString() + ext;
                    fullPath = dir + "/数据导入文件" + ext;
                    context.Request.Files[0].SaveAs(fullPath);

                    file.SaveAs(fullPath);
                    ExcelHelper excel = new ExcelHelper(fullPath);
                    System.Data.DataTable dt = excel.ExcelToDataTable("Sheet1", true);
                    if (dt.Rows.Count > 0)
                    {
                        #region 第③步批量插入
                        //连接数据库
                        try
                        {
                            Demo.Model.UserInfo entity = new Demo.Model.UserInfo();
                            for (int _count = 0; _count < dt.Rows.Count; _count++)
                            {
                                string UserName = dt.Rows[_count]["姓名"].ToString();
                                string Sex = dt.Rows[_count]["性别"].ToString();
                                string Age = dt.Rows[_count]["年龄"].ToString();
                                #region 一级
                                entity.UserName = UserName;
                                entity.Sex = Convert.ToChar(Sex);
                                entity.Age = Convert.ToInt32(Age);
                                string sql = "insert into UserInfo values('" + entity.UserName + "','" + entity.Sex + "','" + entity.Age + "')";
                                bool result = new DBHelper().ExecuteNonQuery(sql);
                                #endregion
                            }
                        }
                        catch (Exception ex)
                        {
                            _json = "{\"state\":\"fail\",\"msg\":\"" + ex.Message.ToString() + "\" }";
                        }
                        #endregion
                    }
                    _json = "{\"state\":\"success\",\"msg\":\"上传成功\" }";
                }
            }
            else
            {
                _json = "{\"state\":\"fail\",\"msg\":\"没有数据\" }";
            }
            return _json;
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
