namespace FF.Common.Excel
{
#if NFX461

using System;
using System.Text;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.Util;

    /// <summary>
    /// Excel读写
    /// </summary>
    public class NPOIExcel : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private ICellStyle cellStyle = null;
        private bool disposed;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="filePath">Excel完整路径</param>
        /// <param name="isNewFile">是否是新建文档  读取模式：false;导出新文件：true。 </param>
        public NPOIExcel(string filePath, bool isNewFile)
        {
            this.fileName = filePath;
            disposed = false;

            //读取模式
            if (isNewFile == false)
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
                workbook = WorkbookFactory.Create(fs);
            }
            else
            {
                fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook();
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook();
                }
            }
        }

        /// <summary>
        /// 读取数据
        /// </summary>
        /// <param name="sheetName">Sheet页名称，空时默认读取第一个Sheet页</param>
        /// <param name="iRowColumnNo">标题行行号</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string sheetName, int iRowColumnNo)
        {
            ISheet sheet = GetWorkSheetBySheetName(sheetName);
            DataTable data = new DataTable();
            int startRow = 0;
            if (sheet != null)
            {
                int cellCount;
                IRow firstRow = sheet.GetRow(iRowColumnNo);
                cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                //读取列表头
                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (cell != null)
                    {
                        string cellValue = cell.ToString();
                        if (cellValue != null)
                        {
                            DataColumn column = new DataColumn(cellValue);
                            data.Columns.Add(column);
                        }
                    }
                }
                startRow = iRowColumnNo + 1;
                //最后一列的标号
                int rowCount = sheet.LastRowNum;
                //读取数据
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　

                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            dataRow[j] = row.GetCell(j).ToString();
                        else {
                            dataRow[j] = null;
                            continue;
                        }
                            
                        //如果是公式Cell 
                        //则仅读取其Cell单元格的显示值 而不是读取公式
                        if (row.GetCell(j).CellType == CellType.Formula)
                        {
                            switch (row.GetCell(j).CellType)
                            {
                                case CellType.Numeric:
                                    dataRow[j] = row.GetCell(j).NumericCellValue;
                                    break;
                                case CellType.String:
                                    dataRow[j] = row.GetCell(j).StringCellValue;
                                    break;
                                case CellType.Boolean:
                                    dataRow[j] = row.GetCell(j).BooleanCellValue;
                                    break;
                                case CellType.Error:
                                    dataRow[j] = row.GetCell(j).ErrorCellValue;
                                    break;
                                default:
                                    try
                                    {
                                        dataRow[j] = row.GetCell(j).NumericCellValue;
                                    }
                                    catch (Exception)
                                    {
                                        dataRow[j] = row.GetCell(j).StringCellValue;
                                    }

                                    break;
                            }
                        }
                        else
                        {

                            switch (row.GetCell(j).CellType)
                            {
                                case CellType.Numeric:
                                    dataRow[j] = row.GetCell(j).NumericCellValue;
                                    break;
                                case CellType.String:
                                    dataRow[j] = row.GetCell(j).StringCellValue;
                                    break;
                                case CellType.Boolean:
                                    dataRow[j] = row.GetCell(j).BooleanCellValue;
                                    break;
                                case CellType.Error:
                                    dataRow[j] = row.GetCell(j).ErrorCellValue;
                                    break;
                                default:
                                    dataRow[j] = row.GetCell(j).ToString();
                                    break;
                            }
                        }
                    }
                    data.Rows.Add(dataRow);
                }
            }
            return data;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="iSheetIndex">要导入的excel的第几个sheet页，默认为0</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, int iSheetIndex, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            try
            {
                if (workbook == null)
                    return -1;
                sheet = GetWorkSheetByIndex(iSheetIndex);



                if (isColumnWritten == true) //写入列名
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
                //写入数据
                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }

                //自适应宽度设置
                for (int columnNum = 0; columnNum <= 26; columnNum++)
                {
                    int columnWidth = sheet.GetColumnWidth(columnNum) / 256;//获取当前列宽度
                    for (int rowNum = 1; rowNum <= sheet.LastRowNum; rowNum++)//在这一列上循环行
                    {
                        IRow currentRow = sheet.GetRow(rowNum);
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.UTF8.GetBytes(currentCell.ToString()).Length;//获取当前单元格的内容宽度
                        if (columnWidth < length + 1)
                        {
                            columnWidth = length + 1;
                        }//若当前单元格内容宽度大于列宽，则调整列宽为当前单元格宽度，后面的+1是我人为的将宽度增加一个字符
                    }
                    sheet.SetColumnWidth(columnNum, columnWidth * 256);
                }
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="FirstRow"></param>
        /// <param name="FirstColumn"></param>
        /// <param name="LastRow"></param>
        /// <param name="LastColumn"></param>
        /// <returns></returns>
        public void MergeCell(int iSheetIndex, int FirstRow, int FirstColumn, int LastRow, int LastColumn)
        {
            ISheet sheet = workbook.GetSheetAt(iSheetIndex);
            CellRangeAddress region = new CellRangeAddress(FirstRow, LastRow, FirstColumn, LastColumn);
            sheet.AddMergedRegion(region);
        }

        /// <summary>
        ///  修改Excel单元格值
        /// </summary>
        /// <param name="iSheetIndex">第几个Sheet页</param>
        /// <param name="strData">修改后的值</param>       
        /// <param name="iRowIndex">行号（从0开始）</param>
        /// <param name="iCellIndex">列号（从0开始）</param>
        /// <returns></returns>
        public bool ModeExcel(int iSheetIndex, string strData, int iRowIndex, int iCellIndex)
        {
            ISheet sheet = workbook.GetSheetAt(iSheetIndex);
            IRow row = sheet.GetRow(iRowIndex);
            ICell Cell = row.GetCell(iCellIndex);
            Cell.SetCellValue(strData);
            return true;
        }

        #region 设置格式
        public void SetCellStyle(int iSheetIndex, int iColumnIndex, string FormatString)
        {
            ISheet sheet = workbook.GetSheetAt(iSheetIndex);
            int row = sheet.LastRowNum;

            for (int i = 0; i < sheet.LastRowNum; i++)
            {
                ICell cell = sheet.GetRow(i).GetCell(iColumnIndex);

                cellStyle = workbook.CreateCellStyle();
                IDataFormat format = workbook.CreateDataFormat();
                cellStyle.DataFormat = format.GetFormat(FormatString);
                cell.CellStyle = cellStyle;
            }
            sheet.ForceFormulaRecalculation = true;
        }
        #endregion

        /// <summary>
        /// 保存文档
        /// </summary>
        public void saveFile()
        {
            workbook.Write(fs); //写入到Excel
            fs.Close();
        }

        /// <summary>
        /// 获取工作Sheet
        /// </summary>
        /// <param name="sheetName">Sheet名称</param>
        /// <returns></returns>
        private ISheet GetWorkSheetBySheetName(string sheetName)
        {
            ISheet sheet = null;
            if (workbook.NumberOfSheets == 0)
            {
                sheet = workbook.CreateSheet();
                return sheet;
            }

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
            return sheet;
        }

        /// <summary>
        /// 获取Sheet页，如果工作薄没有sheet页，则默认创建一个
        /// </summary>
        /// <param name="iSheetIndex"></param>
        /// <returns></returns>
        private ISheet GetWorkSheetByIndex(int iSheetIndex)
        {
            ISheet sheet = null;
            if (workbook.NumberOfSheets == 0)
            {
                sheet = workbook.CreateSheet();
                return sheet;
            }

            sheet = workbook.GetSheetAt(iSheetIndex);
            if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
            {
                throw new Exception("不存在的Sheet页");
            }

            return sheet;
        }

        /// <summary>
        /// 资源释放
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing"></param>
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
    }
#endif 

}
