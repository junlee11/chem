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
using static System.Net.Mime.MediaTypeNames;
using System.Data;
using System.Windows.Media;
using System.IO.Ports;
using static LGchem2.ExcelControl;

namespace LGchem2
{
    public class ExcelControl
    {
        public enum Spec
        {
            spcOut, spcIn, lclOut, nospc
        }

        public void DataTableToExcelQuick(DataTable dt, Workbook workBook, string sheetName, int cell_row, int cell_col, Spec spec, string pdf_path = null)
        {
            if (dt == null) return;            
                      
            Worksheet worksheet = workBook.Worksheets.Item[sheetName];
            worksheet.Columns[1].ColumnWidth = 13;

            //파일명
            if (pdf_path != null)
                worksheet.Cells[cell_row, cell_col] = Path.GetFileName(pdf_path);

            if (sheetName == "SPC")
                worksheet.Columns[2].AutoFit();

            cell_row++;

            //RRT 테이블 Peak 스펙인아웃 판정
            if (spec == Spec.spcOut && pdf_path == null)
            {
                worksheet.Cells[cell_row, cell_col].Offset[5, 1].Font.ColorIndex = 3;
            }
            else if (spec == Spec.lclOut && pdf_path == null)
            {
                worksheet.Cells[cell_row, cell_col].Cells.Offset[5, 1].Font.ColorIndex = 5;
            }

            Range rng = worksheet.Range[worksheet.Cells[cell_row, cell_col], worksheet.Cells[cell_row + dt.Rows.Count, cell_col + dt.Columns.Count - 1]];

            //칼럼부터
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                worksheet.Cells[cell_row, cell_col + j] = dt.Columns[j].ColumnName;
            }
            cell_row++;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    worksheet.Cells[cell_row + i, cell_col + j] = dt.Rows[i][j].ToString();
                }
            }

            this.RangeBorder(rng);
            //if (pdf_path != null) this.ExcelInsertOLE(worksheet, pdf_path, cell_row - 2);
            workBook.Worksheets[1].Activate();
        }

        public void DataTableToExcel(DataTable dt, string path, string sheetName, int cell_row, int cell_col, Spec spec, string pdf_path = null)
        {
            if (dt == null) return;
            Excel.Application application = null;
            Workbook workBook = null;            

            try
            {
                //Excel 프로그램 실행
                application = new Excel.Application();
                //Excel 화면 띄우기 옵션
                application.Visible = false;
                //파일로부터 불러오기
                workBook = application.Workbooks.Open(path);
                //Worksheet worksheet = workBook.ActiveSheet;
                Worksheet worksheet = workBook.Worksheets.Item[sheetName];

                worksheet.Columns[1].ColumnWidth = 13;                

                //파일명
                if (pdf_path != null)                
                    worksheet.Cells[cell_row, cell_col] = Path.GetFileName(pdf_path);

                if (sheetName == "SPC")
                    worksheet.Columns[2].AutoFit();
                    


                cell_row++;

                //RRT 테이블 Peak 스펙인아웃 판정
                if (spec == Spec.spcOut && pdf_path == null)
                {                    
                    worksheet.Cells[cell_row, cell_col].Offset[5, 1].Font.ColorIndex = 3;
                }
                else if (spec == Spec.lclOut && pdf_path == null)
                {
                    worksheet.Cells[cell_row, cell_col].Cells.Offset[5, 1].Font.ColorIndex = 5;
                }

                Range rng = worksheet.Range[worksheet.Cells[cell_row, cell_col], worksheet.Cells[cell_row + dt.Rows.Count, cell_col + dt.Columns.Count - 1]];

                //칼럼부터
                for (int j = 0;j<dt.Columns.Count;j++)
                {
                    worksheet.Cells[cell_row, cell_col + j] = dt.Columns[j].ColumnName;
                }
                cell_row++;

                for (int i = 0;i<dt.Rows.Count;i++)
                {
                    for (int j = 0;j< dt.Columns.Count;j++)
                    {
                        worksheet.Cells[cell_row + i, cell_col + j] = dt.Rows[i][j].ToString();
                    }
                }

                this.RangeBorder(rng);
                if (pdf_path != null) this.ExcelInsertOLE(worksheet, pdf_path, cell_row - 2);
                workBook.Worksheets[1].Activate();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                object missing = Type.Missing;
                object noSave = true;
                workBook.Close(noSave, missing, missing); // 엑셀 워크북 종료
                application.Quit();        // 엑셀 어플리케이션 종료

                //오브젝트 해제
                Global.ReleaseExcelObject(workBook);
                Global.ReleaseExcelObject(application);
            }
        }
        public void SheetNameChange(string path, int pos, string name)
        {
            Excel.Application application = null;
            Workbook workBook = null;

            try
            {
                //Excel 프로그램 실행
                application = new Excel.Application();
                //Excel 화면 띄우기 옵션
                application.Visible = false;
                //파일로부터 불러오기
                workBook = application.Workbooks.Open(path);
                if (workBook.Sheets.Count == 1 && workBook.Sheets.Count == pos)
                {
                    Worksheet worksheet = workBook.Worksheets[pos];
                    worksheet.Name = name;
                }
                else if (workBook.Sheets.Count < pos)
                {
                    //Worksheet worksheet = workBook.Worksheets.Add(Type.Missing, workBook.Worksheets[workBook.Sheets.Count + 1]);
                    Worksheet worksheet = workBook.Worksheets.Add(After: workBook.Sheets[workBook.Sheets.Count]) as Excel.Worksheet;
                    worksheet.Name = name;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                object missing = Type.Missing;
                object noSave = true;
                workBook.Close(noSave, missing, missing); // 엑셀 워크북 종료
                application.Quit();        // 엑셀 어플리케이션 종료

                //오브젝트 해제
                Global.ReleaseExcelObject(workBook);
                Global.ReleaseExcelObject(application);
            }
        }

        public void SaveExcelFile(string path)
        {
            Excel.Application excel = new Excel.Application();
            Workbook workBook = null;
            try
            {                
                excel.Workbooks.Add();
                workBook = excel.ActiveWorkbook;
                Excel.Worksheet sheet = workBook.ActiveSheet;

                //sheet.Cells[1, 1] = path;
                workBook.SaveAs(path);
                workBook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();        // 엑셀 어플리케이션 종료
                //오브젝트 해제
                Global.ReleaseExcelObject(workBook);
                Global.ReleaseExcelObject(excel);
            }
        }

        public void SaveExcelFileQuick(string path)
        {
            Excel.Application excel = new Excel.Application();
            Workbook workBook = null;
            try
            {
                excel.Workbooks.Add();
                workBook = excel.ActiveWorkbook;
                Excel.Worksheet sheet = workBook.ActiveSheet;

                //sheet.Cells[1, 1] = path;
                workBook.SaveAs(path);
                workBook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();        // 엑셀 어플리케이션 종료
                //오브젝트 해제
                Global.ReleaseExcelObject(workBook);
                Global.ReleaseExcelObject(excel);
            }
        }

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

        public void ExcelInsertOLE(Worksheet sheet, string path, int cell_row)
        {
            Excel.OLEObjects oleObjects = (Excel.OLEObjects)
                sheet.OLEObjects(Type.Missing);

            oleObjects.Add(
                Type.Missing,   // ClassType
                path,           // Filename
                false,           // Link
                true,           // DisplayAsIcon
                @"test",   // IconFileName
                0,   // IconIndex
                "PDF File",   // IconLabel
                //55,   // Left
                5,   // Left
                16.5 * cell_row,   // Top
                10,   // Width
                10    // Height
            );
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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

        private void RangeBorder(Range rng)
        {
            rng.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rng.Borders.Weight = Excel.XlBorderWeight.xlThin;
            rng.Interior.ColorIndex = 0;
        }

    }
}
