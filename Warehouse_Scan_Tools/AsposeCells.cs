using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Aspose.Cells;

namespace binSummary
{
    class AsposeCells
    {
        public DataTable ReadExcelNoTitle(string path)
        {
            Workbook workbook = new Workbook(path);
            Cells cells = workbook.Worksheets[0].Cells;
            System.Data.DataTable dataTable = cells.ExportDataTableAsString(1, 0, cells.MaxDataRow, cells.MaxColumn);//没有标题
            //System.Data.DataTable dataTable = cells.ExportDataTable(1, 0, cells.MaxDataRow, cells.MaxColumn);//没有标题
            //System.Data.DataTable dataTable = cells.ExportDataTable(0, 0, cells.MaxDataRow + 1, cells.MaxColumn, true);//有标题
            return dataTable;
        }

        public DataTable ReadExcelTitle(string path)
        {
            Workbook workbook = new Workbook(path);
            Cells cells = workbook.Worksheets[0].Cells;
            System.Data.DataTable dataTable = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow+1, cells.MaxColumn);//有标题
            //System.Data.DataTable dataTable = cells.ExportDataTable(1, 0, cells.MaxDataRow, cells.MaxColumn);//没有标题
            //System.Data.DataTable dataTable = cells.ExportDataTable(0, 0, cells.MaxDataRow + 1, cells.MaxColumn, true);//有标题
            return dataTable;
        }

        /// <summary>
        /// DataTable数据导出Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filepath"></param>
        public static void DataTableExport(DataTable data, string filepath)
        {
            try
            {
                //Workbook book = new Workbook("E:\\test.xlsx"); //打开工作簿
                Workbook book = new Workbook(); //创建工作簿
                Worksheet sheet = book.Worksheets[0]; //创建工作表
                Cells cells = sheet.Cells; //单元格
                //创建样式
                Aspose.Cells.Style style = book.CreateStyle(); ;
                style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 左边界线  
                style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 右边界线  
                style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 上边界线  
                style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 下边界线   
                style.HorizontalAlignment = TextAlignmentType.Center; //单元格内容的水平对齐方式文字居中
                style.Font.Name = "宋体"; //字体
                style.Font.IsBold = true; //设置粗体
                style.Font.Size = 11; //设置字体大小
                //style.ForegroundColor = System.Drawing.Color.FromArgb(153, 204, 0); //背景色
                //style.Pattern = Aspose.Cells.BackgroundType.Solid; //背景样式
                //style.IsTextWrapped = true; //单元格内容自动换行

                //表格填充数据
                int Colnum = data.Columns.Count;//表格列数 
                int Rownum = data.Rows.Count;//表格行数 
                //生成行 列名行 
                for (int i = 0; i < Colnum; i++)
                {
                    cells[0, i].PutValue(data.Columns[i].ColumnName); //添加表头
                    cells[0, i].SetStyle(style); //添加样式
                    //cells.SetColumnWidth(i, data.Columns[i].ColumnName.Length * 2 + 1.5); //自定义列宽
                    //cells.SetRowHeight(0, 30); //自定义高
                }
                //生成数据行 
                for (int i = 0; i < Rownum; i++)
                {
                    for (int k = 0; k < Colnum; k++)
                    {
                        cells[1 + i, k].PutValue(data.Rows[i][k].ToString()); //添加数据
                        cells[1 + i, k].SetStyle(style); //添加样式
                    }                   
                }
                sheet.AutoFitColumns(); //自适应宽
                book.Save(filepath); //保存
                GC.Collect();
            }
            catch (Exception e)
            {
                //logger.Error("生成excel出错：" + e.Message);
            }
        }


        /// <summary>
        /// DataTable数据导出Excel,有特殊格式
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filepath"></param>
        public static void DataTableExportStyle(DataTable data, string filepath)
        {
            try
            {
                //Workbook book = new Workbook("E:\\test.xlsx"); //打开工作簿
                Workbook book = new Workbook(); //创建工作簿
                Worksheet sheet = book.Worksheets[0]; //创建工作表
                Cells cells = sheet.Cells; //单元格
                //创建样式
                Aspose.Cells.Style style = book.CreateStyle(); ;
                style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 左边界线  
                style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 右边界线  
                style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 上边界线  
                style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 下边界线   
                style.HorizontalAlignment = TextAlignmentType.Center; //单元格内容的水平对齐方式文字居中
                style.Font.Name = "等线"; //字体
                //style.Font.IsBold = true; //设置粗体
                style.Font.Size = 11; //设置字体大小
                //style.ForegroundColor = System.Drawing.Color.FromArgb(153, 204, 0); //背景色
                //style.Pattern = Aspose.Cells.BackgroundType.Solid; //背景样式
                //style.IsTextWrapped = true; //单元格内容自动换行

                //表格填充数据
                int Colnum = data.Columns.Count;//表格列数 
                int Rownum = data.Rows.Count;//表格行数 
                //生成行 列名行 
                for (int i = 0; i < Colnum; i++)
                {
                    cells[0, i].PutValue(data.Columns[i].ColumnName); //添加表头
                    cells[0, i].SetStyle(style); //添加样式
                    //cells.SetColumnWidth(i, data.Columns[i].ColumnName.Length * 2 + 1.5); //自定义列宽
                    //cells.SetRowHeight(0, 30); //自定义高
                }
                //生成数据行 
                for (int i = 0; i < Rownum; i++)
                {
                    for (int k = 0; k < Colnum; k++)
                    {
                        cells[1 + i, k].PutValue(data.Rows[i][k].ToString()); //添加数据
                        cells[1 + i, k].SetStyle(style); //添加样式
                    }
                }
                sheet.AutoFitColumns(); //自适应宽
                book.Save(filepath); //保存
                GC.Collect();
            }
            catch (Exception e)
            {
                //logger.Error("生成excel出错：" + e.Message);
            }
        }

        /// <summary>
        /// DataTable数据导出Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filepath"></param>
        /// <param name="sheetNo">多个sheet</param>
        public static void DataTableExport(DataTable[] datas, string filepath,int sheetNo)
        {
            try
            {
                //Workbook book = new Workbook("E:\\test.xlsx"); //打开工作簿
                Workbook book = new Workbook(); //创建工作簿
                book.Worksheets.Clear();
                book.Worksheets.Add("Probecard status");
                book.Worksheets.Add("报废及归还客户针卡");               
                for (int j = 0; j < sheetNo; j++)
                {
                    DataTable data = datas[j];
                    Worksheet sheet = book.Worksheets[j]; //创建工作表
                    Cells cells = sheet.Cells; //单元格
                    //创建样式
                    Aspose.Cells.Style style = book.CreateStyle(); 
                    style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 左边界线  
                    style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 右边界线  
                    style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 上边界线  
                    style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 下边界线   
                    style.HorizontalAlignment = TextAlignmentType.Center; //单元格内容的水平对齐方式文字居中
                    style.VerticalAlignment = TextAlignmentType.Center; 
                    style.Font.Name = "微软雅黑"; //字体
                    style.Font.IsBold = true; //设置粗体
                    style.Font.Size = 12; //设置字体大小
                    //style.ForegroundColor = System.Drawing.Color.FromArgb(153, 204, 0); //背景色
                    //style.Pattern = Aspose.Cells.BackgroundType.Solid; //背景样式
                    style.IsTextWrapped = true; //单元格内容自动换行

                    //表格填充数据
                    int Colnum = data.Columns.Count;//表格列数 
                    int Rownum = data.Rows.Count;//表格行数 

                    if (j == 0)
                    {
                        cells[0, 0].PutValue("ZSW Probe Card"); //添加表头
                       // cells[0, 0].SetStyle(style); //添加样式
                        Range range = cells.CreateRange(0, 0, 1, Colnum);
                        range.SetStyle(style); //添加样式
                        cells.Merge(0, 0, 1, Colnum);//合并单元格
                        cells[0, 0].SetStyle(style); //添加样式
                    }
                    //生成行 列名行 
                    style.Font.Size = 10;
                    style.ForegroundColor = System.Drawing.Color.FromArgb(204, 236, 255);
                    style.Pattern = Aspose.Cells.BackgroundType.Solid;

                    for (int i = 0; i < Colnum; i++)
                    {
                        if (j == 0)
                        {
                            cells[1, i].PutValue(data.Columns[i].ColumnName); //添加表头
                            cells[1, i].SetStyle(style); //添加样式
                            cells.SetRowHeight(0, 20); //自定义高
                            cells.SetRowHeight(1, 30); //自定义高
                        }
                        else
                        {
                            cells[0, i].PutValue(data.Columns[i].ColumnName); //添加表头
                            cells[0, i].SetStyle(style); //添加样式
                            cells.SetRowHeight(0, 30); //自定义高
                        }
                        if (data.Columns[i].ToString().Contains("针卡ID") || data.Columns[i].ToString().Contains("状态") || data.Columns[i].ToString().Equals("Device"))
                        { cells.SetColumnWidth(i, data.Columns[i].ColumnName.Length * 3 + 12); }
                        else
                        {
                            cells.SetColumnWidth(i, data.Columns[i].ColumnName.Length * 2 + 5); //自定义列宽
                        }
                       
                    }
                    //生成数据行
                    style.Font.IsBold = false;
                    style.ForegroundColor = System.Drawing.Color.White;
                    style.Pattern = Aspose.Cells.BackgroundType.Solid;
                    for (int i = 0; i < Rownum; i++)
                    {
                        int o=i;
                        if (j == 0) {  o= i + 1; }
                        for (int k = 0; k < Colnum; k++)
                        {
                            if ((data.Rows[i][k].ToString().Length>0)&(data.Columns[k].ColumnName.Contains("mil") || data.Columns[k].ColumnName.Contains("/") || data.Columns[k].ColumnName.Contains("数")))
                            {
                                 cells[1 + o, k].PutValue(Convert.ToDouble(data.Rows[i][k].ToString())); //添加数据
                            }
                            if ((data.Rows[i][k].ToString().Length > 0) & data.Columns[k].ColumnName.Contains("Device"))
                            {
                                cells[1 + o, k].PutValue(data.Rows[i][k].ToString().Replace("，", ",").Replace(";", ",").Replace(",", "\n")); //添加数据
                            }
                            else
                            {
                                cells[1 + o, k].PutValue(data.Rows[i][k].ToString()); //添加数据
                            }
                            cells[1 + o, k].SetStyle(style); //添加样式
                        }
                    }
                    if (j == 0)
                    {
                        cells[Rownum + 2, 0].PutValue("更新人员：ATEC  更新日期："+DateTime.Now.Date.ToShortDateString()); //添加表头
                        //cells[Rownum + 2, 0].SetStyle(style); //添加样式
                        Range range = cells.CreateRange(Rownum + 2, 0, 1, Colnum);
                        range.SetStyle(style); //添加样式
                        cells.Merge(Rownum + 2, 0, 1,Colnum);//合并单元格
                       // cells[Rownum + 2, 0].SetStyle(style);
                    }
                   // sheet.AutoFitColumns(); //自适应宽
                }
                book.Save(filepath); //保存
                GC.Collect();
            }
            catch (Exception e)
            {
                //logger.Error("生成excel出错：" + e.Message);
            }
        }

        /// <summary>
        /// DataTable数据导出Excel,有特殊格式
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filepath"></param>
        public static void ExportBinSummary(DataTable data, string filepath)
        {
            try
            {
                //Workbook book = new Workbook("E:\\test.xlsx"); //打开工作簿
                Workbook book = new Workbook(); //创建工作簿
                Worksheet sheet = book.Worksheets[0]; //创建工作表
                Cells cells = sheet.Cells; //单元格
                //创建样式
                Aspose.Cells.Style style = book.CreateStyle(); ;
                style.Borders[Aspose.Cells.BorderType.LeftBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 左边界线  
                style.Borders[Aspose.Cells.BorderType.RightBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 右边界线  
                style.Borders[Aspose.Cells.BorderType.TopBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 上边界线  
                style.Borders[Aspose.Cells.BorderType.BottomBorder].LineStyle = Aspose.Cells.CellBorderType.Thin; //应用边界线 下边界线   
                style.HorizontalAlignment = TextAlignmentType.Center; //单元格内容的水平对齐方式文字居中
                style.Font.Name = "等线"; //字体
                //style.Font.IsBold = true; //设置粗体
                style.Font.Size = 11; //设置字体大小
                //style.ForegroundColor = System.Drawing.Color.FromArgb(153, 204, 0); //背景色
                //style.Pattern = Aspose.Cells.BackgroundType.Solid; //背景样式
                //style.IsTextWrapped = true; //单元格内容自动换行

                //表格填充数据
                int Colnum = data.Columns.Count;//表格列数 
                int Rownum = data.Rows.Count;//表格行数 
                //生成行 列名行 
                for (int i = 0; i < Colnum; i++)
                {
                    cells[0, i].PutValue(data.Columns[i].ColumnName); //添加表头
                    cells[0, i].SetStyle(style); //添加样式
                    //cells.SetColumnWidth(i, data.Columns[i].ColumnName.Length * 2 + 1.5); //自定义列宽
                    //cells.SetRowHeight(0, 30); //自定义高
                }
                //生成数据行 
                for (int i = 0; i < Rownum; i++)
                {
                    for (int k = 0; k < Colnum; k++)
                    {
                        if ((data.Rows[i][k].ToString().Length > 0) & (data.Columns[k].ColumnName.Contains("Bin") || data.Columns[k].ColumnName.Contains("Total")))
                        {
                            cells[1 + i, k].PutValue(Convert.ToInt32(data.Rows[i][k].ToString())); //添加数据
                        }
                        else
                        {
                            cells[1 + i, k].PutValue(data.Rows[i][k].ToString()); //添加数据
                        }
                        cells[1 + i, k].SetStyle(style); //添加样式
                    }
                }
                sheet.AutoFitColumns(); //自适应宽
                book.Save(filepath); //保存
                GC.Collect();
            }
            catch (Exception e)
            {
                //logger.Error("生成excel出错：" + e.Message);
            }
        }

    }
}
