using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace LGchem2
{
    public class ExcelControl
    {
        public (string, DataTable) GetDic_SheetContentTable(string ref_path, string mat_name)
        {
            string key = "";
            string val_str = "";
            
            (key, val_str) = this.GetDic_SheetContentTxt(ref_path, mat_name);
            DataTable dt_rst = new DataTable();
            if (key != "")
            {
                dt_rst = this.GetCSVDataFromText(val_str);                
                dt_rst = Global.DelLittleRow(dt_rst, 0);
            }                

            return (key, dt_rst);
        }

        public void ExcelInsertOLE(string path)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Workbook workBook = excel.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet sheet = workBook.ActiveSheet;

            Excel.OLEObjects oleObjects = (Microsoft.Office.Interop.Excel.OLEObjects)
                sheet.OLEObjects(Type.Missing);

            oleObjects.Add(
                Type.Missing,   // ClassType
                path,           // Filename
                false,           // Link
                true,           // DisplayAsIcon
                @"C:\Users\USER\Desktop\코드\a.ico",   // IconFileName
                0,   // IconIndex
                Path.GetFileName(path),   // IconLabel
                50,   // Left
                50,   // Top
                50,   // Width
                50    // Height
            );

            excel.Visible = true;
            workBook.Close(true);
            excel.Quit();
        }

        //private
        private (string, string) GetDic_SheetContentTxt(string ref_path, string mat_name)
        {
            string key = "";
            string val = "";

            Microsoft.Office.Interop.Excel.Application application = null;
            Workbook workBook = null;
            List<string> buff = new List<string>();

            try
            {
                //Excel 프로그램 실행
                application = new Microsoft.Office.Interop.Excel.Application();
                //Excel 화면 띄우기 옵션
                application.Visible = false;
                //파일로부터 불러오기
                workBook = application.Workbooks.Open(ref_path);

                // Sheet항목들을 돌아가면서 내용을 확인
                foreach (Excel.Worksheet workSheet in workBook.Worksheets)
                {
                    //골뱅이 앞 파일명에 워크시트 이름이 포함되는 경우만 레퍼런스
                    if (!mat_name.Contains(workSheet.Name)) continue;

                    Excel.Range range = workSheet.UsedRange;    // 사용중인 셀 범위를 가져오기
                    string temp_rst = "";
                    // 가져온 행(row) 만큼 반복
                    for (int row = 1; row <= range.Rows.Count; row++)
                    {
                        List<string> lstCell = new List<string>();

                        // 가져온 열(row) 만큼 반복
                        for (int column = 1; column <= range.Columns.Count; column++)
                        {
                            string str = "";
                            object obj = (range.Cells[row, column] as Excel.Range).Value2;
                            if (obj != null) str = obj.ToString();  // 셀 데이터 가져옴
                            lstCell.Add(str); // 리스트에 할당
                        }
                        if (row == 1) temp_rst = string.Join(",", lstCell.ToArray());
                        else temp_rst = temp_rst + "\n" + string.Join(",", lstCell.ToArray());

                        //buff.Add(string.Join(",", lstCell.ToArray())); // 표시용 데이터 추가
                    }
                    key = workSheet.Name;
                    val = temp_rst;                    
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
            finally
            {
                object missing = Type.Missing;
                object noSave = false;
                workBook.Close(noSave, missing, missing); // 엑셀 웨크북 종료
                application.Quit();        // 엑셀 어플리케이션 종료

                //오브젝트 해제
                Global.ReleaseExcelObject(workBook);
                Global.ReleaseExcelObject(application);
            }

            return (key, val);
        }
        
        private DataTable GetCSVDataFromText(string txt)            
        {
            try
            {
                string[] arr_txt = txt.Split('\n');

                DataTable table = new DataTable();
                var flag_dtcolumn = 0;

                foreach (string line in arr_txt) 
                {                    
                    string[] data = line.Split(',');

                    if (flag_dtcolumn == 0)
                    {
                        foreach (string s in data)
                        {
                            table.Columns.Add(s);
                        }
                    }
                    else
                    {
                        table.Rows.Add(data.ToArray());
                    }

                    flag_dtcolumn++;

                }

                return table;
            }
            catch
            {
                return null;
            }

        }

    }
}
