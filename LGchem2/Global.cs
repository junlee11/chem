using Aspose.Pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace LGchem2
{
    public class Global
    {
        public static bool DatatableToCSV(DataTable dtDataTable, string strFilePath)
        {
            try
            {
                StreamWriter sw = new StreamWriter(strFilePath, false);
                //headers    
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    sw.Write(dtDataTable.Columns[i]);
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
                int br_flag = 0;
                //this.print_DataTable(dtDataTable);
                foreach (DataRow dr in dtDataTable.Rows)
                {
                    for (int i = 0; i < dtDataTable.Columns.Count; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            string value = dr[0].ToString();
                            if (i == 0 && value == "")
                            {
                                br_flag = 1;
                                break;
                            }
                            if (value.Contains(','))
                            {
                                value = String.Format("\"{0}\"", value);
                                Debug.WriteLine(value);
                                sw.Write(value);
                            }
                            else
                            {
                                sw.Write(dr[i].ToString());
                            }
                        }
                        if (i < dtDataTable.Columns.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    if (br_flag == 1) break;

                    sw.Write(sw.NewLine);
                }
                sw.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("데이터테이블을 csv로 저장할 수 없습니다." + "\n" + ex.Message);
                return false;
            }
        }

        public static DataTable GetCSVData(string str_path, int resource_flag)            //2이면 임베디드 리소스
        {
            try
            {
                StreamReader file;
                if (resource_flag == 1) file = new StreamReader(str_path);
                else
                {
                    var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(str_path);
                    file = new StreamReader(stream);
                }

                DataTable table = new DataTable();
                var flag_dtcolumn = 0;

                while (!file.EndOfStream)
                {
                    string line = file.ReadLine();
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

        

        public static void print_DataTable(DataTable dt, int header = 1)
        {
            // header = 1 -> Print header, header = 0 -> No print header
            int m = dt.Columns.Count; // number of column
            int n = dt.Rows.Count; // number of row

            string[] line = new string[m];
            string[] result = new string[n + header];
            if (header == 1)
            {
                //result[0] = "\t" + String.Join("\t", names(dt));
                List<String> list_col = new List<String>();
                foreach (DataColumn col in dt.Columns)
                {
                    list_col.Add(col.ColumnName);
                }
                result[0] = "\t" + String.Join("\t", list_col);

            }
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m; j++)
                    line[j] = dt.Rows[i][j].ToString();
                result[i + header] = i + "\t" + String.Join("\t", line);
            }
            foreach (var e in result)
            {
                Debug.WriteLine(e);
            }
        }

        public static void AddColDt(DataTable dt, string col, int cnt)
        {
            if (dt.Columns.Contains(col))
            {
                cnt++;
                col = $"{col}_{cnt.ToString()}";
                AddColDt(dt, col, cnt);
            }
            else
            {
                //신규 칼럼
                DataColumn dataColumn = new DataColumn();
                dataColumn.ColumnName = col;
                dt.Columns.Add(dataColumn);
            }
        }

        public static DataTable DelEmptyColumn(DataTable dt)
        {
            foreach (var column in dt.Columns.Cast<DataColumn>().ToArray())
            {
                if (dt.AsEnumerable().All(dr => dr.IsNull(column)))
                //if (dt.AsEnumerable().All(dr => dr.ToString().Trim() == ""))
                    dt.Columns.Remove(column);
            }

            return dt;
        }

        public static DataTable DelLittleRow(DataTable dt, int min_cnt)
        {
            dt.AcceptChanges();
            foreach (DataRow dr in dt.Rows)
            {
                int cnt = 0;
                //데이터가 있는열이 1개 미만인 행은 제거
                for (int i =0;i<dt.Columns.Count;i++)
                {
                    if (dr[i].ToString() != "") cnt++;
                }
                if (cnt <= min_cnt) dr.Delete();
            }
            dt.AcceptChanges();

            return dt;
        }

        public static double? VlookupDt(DataTable dt, double ref_val, string ref_col, string find_col)
        {
            foreach (DataRow dr in dt.Rows)
            {
                if (ref_val == double.Parse(dr[ref_col].ToString())) return double.Parse(dr[find_col].ToString());
            }
            return null;
        }

        public static int? VlookupDt_Int(DataTable dt, double ref_val, string ref_col, string find_col)
        {
            foreach (DataRow dr in dt.Rows)
            {
                if (ref_val == double.Parse(dr[ref_col].ToString())) return Int32.Parse(dr[find_col].ToString());
            }
            return null;
        }

        public static bool ChkStrInDicKey(string str, Dictionary<string, DataTable> dic)
        {
            foreach (KeyValuePair<string, DataTable> items in dic)
            {
                if (str.Contains(items.Key)) return true;
            }
            return false;
        }
        public static DataTable GetdtInDicKey(string str, Dictionary<string, DataTable> dic)
        {
            foreach (KeyValuePair<string, DataTable> items in dic)
            {
                if (str.Contains(items.Key)) return items.Value;
            }
            return null;
        }
        
        public static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
