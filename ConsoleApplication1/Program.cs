using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;

using Microsoft.Office.Interop.Excel;

namespace test1
{
    class Program
    {
        public static Array ReadXls(string filename, int index)//讀取第index個sheet的資料
        {
            //啟動Excel應用程式
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            //開啟filename表
            _Workbook book = xls.Workbooks.Open(filename);

            _Worksheet sheet;//定義sheet變數
            xls.Visible = false;//設定Excel後臺執行
            xls.DisplayAlerts = false;//設定不顯示確認修改提示

            try
            {
                sheet = (_Worksheet)book.Worksheets.get_Item(index);//獲得第index個sheet，準備讀取
            }
            catch (Exception ex)//不存在就退出
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            Console.WriteLine(sheet.Name);
            int row = sheet.UsedRange.Rows.Count;//獲取不為空的行數
            int col = sheet.UsedRange.Columns.Count;//獲取不為空的列數

            // Array value = (Array)sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[row, col]).Cells.Value2;//獲得區域資料賦值給Array陣列，方便讀取

            Microsoft.Office.Interop.Excel.Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[row, col]];
            Array value = (Array)range.Value2;

            book.Save();//儲存
            book.Close(false, Missing.Value, Missing.Value);//關閉開啟的表
            xls.Quit();//Excel程式退出
            //sheet,book,xls設定為null，防止記憶體洩露
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();//系統回收資源
            return value;
        }

        static void Main(string[] args)
        {
            string Current;
            Current = Directory.GetCurrentDirectory();//獲取當前根目錄
            //===== 紀錄時間 =====
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw.Reset();//碼表歸零
            sw.Start();//碼表開始計時

            Array Data = ReadXls(Current + "\\AB_TEMPA.csv", 1);//讀取test.xlsx的第一個sheet表
            sw.Stop();//碼錶停止

            //印出所花費的總豪秒數
            Console.WriteLine(sw.Elapsed.TotalMilliseconds.ToString());

            //===== 紀錄時間 =====

            /*
            foreach (var temp in Data)
            {
                Console.WriteLine(temp);
            }
            */
            Console.ReadKey();
        }
    }
}