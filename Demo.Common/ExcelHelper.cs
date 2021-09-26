using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Data;
using System.IO;
using System;
using NPOI.SS.Util;

namespace Demo.Common
{
    public class ExcelHelper
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;
        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            disposed = false;
        }
        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();
            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }
                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }
                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }
        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);
                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数
                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }
                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }
                fs = null;
                disposed = true;
            }
        }

        /*----------------------------2021.5.20 16:51--------------------------------*/
        public void ExcelOperationClass()
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
        }
    }
}
