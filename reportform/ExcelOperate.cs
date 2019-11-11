using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
namespace reportform
{
  public static class Excel
  {
    /// <summary>
    /// 从EXCEL读取数据
    /// </summary>
    /// <returns></returns>
    public static DataTable ReadExcel()
    {
      string path = AppDomain.CurrentDomain.BaseDirectory;
      DataTable dt = new DataTable();
      XSSFWorkbook wb = new XSSFWorkbook(path + "Tag.xlsx");
      XSSFSheet ws = (XSSFSheet)wb.GetSheetAt(0);//表
      DataColumn[] dc = new DataColumn[wb[0].GetRow(0).LastCellNum + 1];
      int rowCount = ws.PhysicalNumberOfRows;
      int rowNums = 0;
      for (int i = 0; i < rowCount; i++)
      {
        var a = ws.GetRow(i);
        if (ws.GetRow(i) != null && ws.GetRow(i).GetCell(ws.GetRow(i).FirstCellNum).ToString() != "")
        {
          rowNums++;
        }
      }
      for (int ColumnNum = 0; ColumnNum < ws.GetRow(0).LastCellNum; ColumnNum++)
      {
        dc[ColumnNum] = new DataColumn(ws.GetRow(0).GetCell(ColumnNum).ToString(), Type.GetType("System.String"));
        if (dc[ColumnNum] != null)
        {
          dt.Columns.Add(dc[ColumnNum]);
        }
      }
      for (int RowNum = 1; RowNum <= rowNums; RowNum++)
      {
        DataRow dr = dt.NewRow();
        for (int ColumnNum = 0; ColumnNum < ws.GetRow(0).LastCellNum; ColumnNum++)
        {
          if (dr[ColumnNum] != null)
          {
            dr[ColumnNum] = ws.GetRow(RowNum).GetCell(ColumnNum).ToString();
          }
        }
        if (dr != null)
        {
          dt.Rows.Add(dr);
        }
      }
      return dt;
    }

    /// <summary>
    /// 根据文件名读取excel文件
    /// </summary>
    /// <param name="fileFullPath">文件名(绝对路径)</param>
    /// <param name="sheetnum">表序号(从0开始)</param>
    /// <returns></returns>
    public static DataTable ReadExcel(string fileFullPath, int sheetnum)
    {
      DataTable dt = new DataTable();
      XSSFWorkbook wb = new XSSFWorkbook(fileFullPath);//文件
      XSSFSheet ws = (XSSFSheet)wb.GetSheetAt(sheetnum);//表
      DataColumn[] dc = new DataColumn[ws.GetRow(0).LastCellNum + 1];
      int rowCount = ws.PhysicalNumberOfRows;
      int rowNums = 0;
      for (int i = 0; i < rowCount; i++)
      {
        var a = ws.GetRow(i);
        if (ws.GetRow(i) != null && ws.GetRow(i).GetCell(ws.GetRow(i).FirstCellNum).ToString() != "")
        {
          rowNums++;
        }
      }
      for (int ColumnNum = 0; ColumnNum < ws.GetRow(0).LastCellNum; ColumnNum++)
      {
        dc[ColumnNum] = new DataColumn(ws.GetRow(0).GetCell(ColumnNum).ToString(), Type.GetType("System.String"));
        if (dc[ColumnNum] != null)
        {
          dt.Columns.Add(dc[ColumnNum]);
        }
      }
      for (int RowNum = 1; RowNum < rowNums; RowNum++)
      {
        DataRow dr = dt.NewRow();
        for (int ColumnNum = 0; ColumnNum < ws.GetRow(0).LastCellNum; ColumnNum++)
        {
          if (dr[ColumnNum] != null)
          {
            var cellValue = ws.GetRow(RowNum).GetCell(ColumnNum);
            dr[ColumnNum] = cellValue == null ? "" : cellValue.ToString();
          }
        }
        if (dr != null)
        {
          dt.Rows.Add(dr);
        }
      }
      return dt;
    }

    /// <summary>
    /// 根据文件名读取excel文件
    /// </summary>
    /// <param name="fileFullPath">文件名(绝对路径)</param>
    /// <param name="sheetname">表名</param>
    /// <returns></returns>
    public static DataTable ReadExcel(string fileFullPath, string sheetname)
    {
      DataTable dt = new DataTable();
      XSSFWorkbook wb = new XSSFWorkbook(fileFullPath);
      XSSFSheet ws = (XSSFSheet)wb.GetSheetAt(wb.GetSheetIndex(sheetname));
      DataColumn[] dc = new DataColumn[ws.GetRow(0).LastCellNum + 1];
      int rowCount = ws.PhysicalNumberOfRows;
      int rowNums = 0;
      for (int i = 0; i < rowCount; i++)
      {
        var a = ws.GetRow(i);
        if (ws.GetRow(i) != null && ws.GetRow(i).GetCell(ws.GetRow(i).FirstCellNum).ToString() != "")
        {
          rowNums++;
        }
      }
      for (int ColumnNum = 0; ColumnNum < ws.GetRow(0).LastCellNum; ColumnNum++)
      {
        dc[ColumnNum] = new DataColumn(ws.GetRow(0).GetCell(ColumnNum).ToString(), Type.GetType("System.String"));
        if (dc[ColumnNum] != null)
        {
          dt.Columns.Add(dc[ColumnNum]);
        }
      }
      for (int RowNum = 1; RowNum < rowNums; RowNum++)
      {
        DataRow dr = dt.NewRow();
        for (int ColumnNum = 0; ColumnNum < ws.GetRow(0).LastCellNum; ColumnNum++)
        {
          if (dr[ColumnNum] != null)
          {
            var cellValue = ws.GetRow(RowNum).GetCell(ColumnNum);
            dr[ColumnNum] = cellValue == null ? "" : cellValue.ToString();
          }
        }
        if (dr != null)
        {
          dt.Rows.Add(dr);
        }
      }
      return dt;
    }

    /// <summary>
    /// 将数据写入excel
    /// </summary>
    /// <param name="filename">文件名</param>
    /// <param name="dt">要写入的datatable</param>
    public static void WriteSheet(string filename, DataTable dt)
    {
      XSSFWorkbook wb = new XSSFWorkbook();//Path + filename);
      int SheetIndex = wb.GetSheetIndex(DateTime.Now.ToString("MMdd"));
      XSSFSheet ws = new XSSFSheet();
      if (SheetIndex != -1)
      {
        ws = (XSSFSheet)wb.GetSheetAt(wb.GetSheetIndex(DateTime.Now.ToString("MMdd")));
      }
      else
      {
        wb.CreateSheet(DateTime.Now.ToString("MMdd"));
        ws = (XSSFSheet)wb.GetSheetAt(wb.GetSheetIndex(DateTime.Now.ToString("MMdd")));
      }
      ws.CreateRow(0);
      for (int i = 0; i < dt.Columns.Count; i++)
      {
        ws.GetRow(0).CreateCell(i);
        ws.GetRow(0).Cells[i].SetCellValue(dt.Columns[i].ColumnName);
      }
      int rownum = 1;
      foreach (DataRow row in dt.Rows)
      {
        try
        {
          ws.CreateRow(rownum);
          for (int i = 0; i < dt.Columns.Count; i++)
          {
            ws.GetRow(rownum).CreateCell(i);
            var value = row[i];
            if (value.GetType() == typeof(Int64))
            {
              ws.GetRow(rownum).Cells[i].SetCellValue((Int64)value);
            }
            else if (value.GetType() == typeof(Double))
            {
              ws.GetRow(rownum).Cells[i].SetCellValue((Double)value);
            }
            else if (value.GetType() == typeof(Boolean))
            {
              ws.GetRow(rownum).Cells[i].SetCellValue((Boolean)value);
            }
            else if (value.GetType() == typeof(DateTime))
            {
              ws.GetRow(rownum).Cells[i].SetCellValue(((DateTime)value).ToString("yyyy-MM-dd"));
            }
            else
            {
              ws.GetRow(rownum).Cells[i].SetCellValue(value.ToString());
            }
          }
          rownum++;
        }
        catch (Exception ex)
        {
                    throw ex;
          //Log.AppendAllText(ex.ToString());
        }
      }
      //写入并保存excel文件，否则打开时会报错
      FileStream file = new FileStream(filename, FileMode.Create);
      wb.Write(file);
      file.Close();
      wb.Close();
    }

    /// <summary>
    /// 从excel读取shape内容
    /// </summary>
    /// <param name="filename"></param>
    /// <param name="sheetname"></param>
    /// <param name="textboxname"></param>
    public static void ReadShape(string fileFullPath, string sheetName)
    {
      DataTable dt = new DataTable();
      XSSFWorkbook wb = new XSSFWorkbook(fileFullPath);//文件
      XSSFSheet ws = (XSSFSheet)wb.GetSheet(sheetName);//表
      XSSFDrawing pat = ws.GetDrawingPatriarch();
      List<XSSFShape> shapes = pat.GetShapes();
      foreach (XSSFShape shape in shapes)
      {
        var a = shape;
        if (shape.GetType() == typeof(XSSFSimpleShape))
        {
          string textbox = ((XSSFSimpleShape)shape).Text;
        }
      }
    }
  }
}
