using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
    }
}
