using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using System.Runtime.InteropServices;



namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        private void btn_Office_Click(object sender, EventArgs e)
         {
             //string importExcelPath = "E:\\import.xlsx";
             string importExcelPath = "C: \\Users\\wuhongjun\\source\\repos\\ExcelAddIn2\\import.xlsx"; 
             //string exportExcelPath = "E:\\export.xlsx";
            string exportExcelPath = "C: \\Users\\wuhongjun\\source\\repos\\ExcelAddIn2\\export.xlsx";
            //创建
            Excel.Application xlApp = new Excel.Application();
             xlApp.DisplayAlerts = false;
             xlApp.Visible = false;
             xlApp.ScreenUpdating = false;
             //打开Excel
             Excel.Workbook xlsWorkBook = xlApp.Workbooks.Open(importExcelPath, System.Type.Missing, System.Type.Missing, System.Type.Missing,
             System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
             System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
 
             //处理数据过程，更多操作方法自行百度
             Excel.Worksheet sheet = xlsWorkBook.Worksheets[1];//工作薄从1开始，不是0
             sheet.Cells[1, 1] = "test";
 
             //另存
             //xlsWorkBook.SaveAs(exportExcelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
             xlsWorkBook.SaveAs(exportExcelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //关闭Excel进程
            ClosePro(xlApp, xlsWorkBook);
         }
 
         public void ClosePro(Excel.Application xlApp, Excel.Workbook xlsWorkBook)
         {
             if (xlsWorkBook != null)
                 xlsWorkBook.Close(true, Type.Missing, Type.Missing);
             xlApp.Quit();
             // 安全回收进程
             System.GC.GetGeneration(xlApp);
             IntPtr t = new IntPtr(xlApp.Hwnd);   //获取句柄
             int k = 0;
             GetWindowThreadProcessId(t, out k);   //获取进程唯一标志
             //--using System.Runtime.InteropServices;
             //--[DllImport("User32.dll", CharSet = CharSet.Auto)]
             //--public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
             //--以上--3句才能保证---GetWindowThreadProcessId(t, out k)函数的正确；




            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
             p.Kill();     //关闭进程
         }



#region VSTO 生成的代码

/// <summary>
/// 设计器支持所需的方法 - 不要修改
/// 使用代码编辑器修改此方法的内容。
/// </summary>
private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
