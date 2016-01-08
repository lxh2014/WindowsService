/* 
 * Copyright (c) 2013，武漢聯綿信息技術有限公司
 * All rights reserved. 
 *  
 * 文件名称：FileManagement 
 * 摘   要： 
 *  
 * 当前版本：1.0 
 * 作   者：lhInc 
 * 完成日期：2013/8/16 13:59:43 
 * 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Web;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Net;

namespace Common
{
    public class NPOI4ExcelHelper
    {
        private IWorkbook hssfworkbook = null;
        private string filePath;
        public NPOI4ExcelHelper(string filePath)
        {
            this.filePath = filePath;
            Init();
        }

        private void Init()
        {
            #region//初始化信息
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (filePath.Trim().ToLower().EndsWith(".xls"))
                    {
                        hssfworkbook = new HSSFWorkbook(file);
                    }
                    else if (filePath.Trim().ToLower().EndsWith(".xlsx"))
                    {
                        hssfworkbook = new XSSFWorkbook(file);
                    }
                    else if (filePath.Trim().ToLower().EndsWith(".csv"))
                    {
                        hssfworkbook = new XSSFWorkbook(file);
                    }
                   
                   
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            #endregion
        }

        public ArrayList SheetNameForExcel()
        {
            if (hssfworkbook == null) { return null; };

            ArrayList names = new ArrayList();
            int i = hssfworkbook.NumberOfSheets;
            for (int j = 0; j < i; j++)
            {
                string sheet = hssfworkbook.GetSheetName(j);
                if (!names.Contains(sheet))
                {
                    names.Add(sheet);
                }
            }
            return names;
        }

        public DataTable SheetData(int sheetIndex = 0)
        {
            if (filePath.Trim().ToLower().EndsWith(".xls"))
            {
                return ExcelToTableForXLS(sheetIndex);
            }
            else if (filePath.Trim().ToLower().EndsWith(".xlsx"))
            {
                return ExcelToTableForXLSX(sheetIndex);
            }
            else if (filePath.Trim().ToLower().EndsWith(".csv"))
            {
                return ExcelToTableForXLSX(sheetIndex);
            }
           
            return null;
        }

        /// <summary>
        /// 讀取excel文件第一個sheet的資料  07
        /// </summary>
        /// <param name="filePath">文件路徑</param>
        /// <returns></returns>
        private DataTable ExcelToTableForXLSX(int sheetIndex = 0)
        {
            if (hssfworkbook == null) { return null; };

            ISheet sheet = hssfworkbook.GetSheetAt(sheetIndex);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable();

            //第一行添加未table的列名
            //一行最后一个方格的编号 即总的列数
            rows.MoveNext();
            IRow headerRow = (XSSFRow)rows.Current;
            for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
            {
                dt.Columns.Add(headerRow.GetCell(j) != null ? headerRow.GetCell(j).StringCellValue.Trim() : "");
            }

            while (rows.MoveNext())
            {
                IRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < (row.LastCellNum <= dt.Columns.Count ? row.LastCellNum : dt.Columns.Count); i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        switch (cell.CellType) { 
                            case CellType.NUMERIC:
                                DateTime value = new DateTime();
                                if (DateTime.TryParse(cell.ToString(), out value))
                                    dr[i] = cell.DateCellValue.ToString("yyyy/MM/dd hh:mm:ss");
                                else
                                    dr[i] = cell.NumericCellValue.ToString();
                                break;
                            default:
                                dr[i] = cell.ToString().Trim();
                                break;
                        }
                        
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
        /// <summary>
        /// 讀取excel文件第一個sheet的資料  03
        /// </summary>
        /// <param name="filePath">文件路徑</param>
        /// <returns></returns>
        private DataTable ExcelToTableForXLS(int sheetIndex = 0)
        {
            if (hssfworkbook == null) { return null; };

            ISheet sheet = hssfworkbook.GetSheetAt(sheetIndex);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable();

            //第一行添加未table的列名
            //一行最后一个方格的编号 即总的列数
            rows.MoveNext();
            IRow headerRow = (HSSFRow)rows.Current;
            for (int j = 0; j < (sheet.GetRow(0).LastCellNum); j++)
            {
                dt.Columns.Add(headerRow.GetCell(j) != null ? headerRow.GetCell(j).StringCellValue.Trim() : "");
            }

            while (rows.MoveNext())
            {
                IRow row = (HSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < (row.LastCellNum <= dt.Columns.Count ? row.LastCellNum : dt.Columns.Count); i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        switch (cell.CellType)
                        {
                            case CellType.NUMERIC:
                                DateTime value = new DateTime();
                                if (DateTime.TryParseExact(cell.ToString(),"M/d/y H:m",null,System.Globalization.DateTimeStyles.None, out value)
                                    ||DateTime.TryParse(cell.ToString(),out value))
                                    dr[i] = cell.DateCellValue.ToString("yyyy/MM/dd hh:mm:ss");
                                else
                                    dr[i] = cell.ToString();
                                break;
                            default:
                                dr[i] = cell.ToString().Trim();
                                break;
                        }
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }

    public class FileManagement
    {
        public FileManagement()
        {

        }

        /// <summary>
        /// 返回連接excel字符串
        /// </summary>
        /// <param name="fileName">數據源地址</param>
        /// <returns></returns>
        public string GetExcelConnStr(string fileName)
        {
            string strConOLE = string.Empty;
            if (!System.IO.File.Exists(fileName))
            {
                return strConOLE;
            }
            string extension = System.IO.Path.GetExtension(fileName).ToLower().ToString();
            if (extension.CompareTo(".xls") == 0)//针对2003格式进行的判断
            {
                strConOLE = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'";
            }
            else if (extension.CompareTo(".xlsx") == 0)//针对2007格式进行的判断
            {
                strConOLE = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
            }
            return strConOLE;
        }

        /// <summary>
        /// 返回Excel 所有sheet名稱
        /// </summary>
        /// <param name="fileName">連接字符串</param>
        /// <returns></returns>
        public ArrayList GetExcelTables(string connStr)
        {
            OleDbConnection oleConn = new OleDbConnection(connStr);
            ArrayList sheetList = new ArrayList();
            try
            {
                if (oleConn.State == ConnectionState.Closed)
                {
                    oleConn.Open();
                }
                DataTable dt_Sheets = oleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                foreach (DataRow row in dt_Sheets.Rows)
                {
                    sheetList.Add(row["TABLE_NAME"].ToString());
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                oleConn.Close();
            }
            return sheetList;
        }

        /// <summary>
        /// 上傳文件
        /// </summary>
        /// <param name="httpPostedFile">源文件</param>
        /// <param name="fileName">保存的文件名</param>
        /// <param name="extensions">允許上傳格式</param>
        /// <param name="maxSize">允許文件上傳最大值 單位KB</param>
        /// <param name="minSize">語序上傳最小值 單位KB</param>
        /// <returns></returns>
        public bool UpLoadFile(HttpPostedFileBase httpPostedFile, string fileName, string extensions, int maxSize, int minSize)
        {
            WebClient wc = new WebClient();
            int fileSize = httpPostedFile.ContentLength;
            string extension = System.IO.Path.GetExtension(fileName).ToLower().ToString();
            extension = extension.Remove(extension.LastIndexOf("."), 1);
            string[] types = extensions.ToLower().Split(',');
            if (extensions.ToLower().IndexOf(extension) < 0)
            {
                return false;
            }
            try
            {
                if (fileSize > maxSize * 1024)
                {
                    return false;
                }
                if (fileSize < minSize * 1024)
                {
                    return false;
                }
                httpPostedFile.SaveAs(fileName);

                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

     

    
        public string NewFileName(string oldName)
        {
            string newName = string.Empty;
            if (oldName.LastIndexOf(".") != -1)
            {
                string[] strs = oldName.Split('.');
                newName = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "." + strs[strs.Length - 1];
            }
            return newName;
        }
    }
}
