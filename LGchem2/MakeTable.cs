using Aspose.Pdf;
using Aspose.Pdf.Plugins;
using Aspose.Pdf.Text;
using javax.smartcardio;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text; 
using System.Threading.Tasks;
using System.Xml.Linq;

namespace LGchem2
{
    public interface ITable
    {
        DataTable Extract_Table(string path);        
    }

    public class MakeRawTable_LGD : ITable
    {
        public const double width_A4 = 597.6;
        public const double height_A4 = 842.4;

        public DataTable Extract_Table(string path)
        {   
            DataTable dt_rst = new DataTable();            
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(path);

            //Debug.WriteLine(pdfDocument.Pages.Count);
            foreach (var page in pdfDocument.Pages)
            {
                Aspose.Pdf.Text.TableAbsorber absorber = new Aspose.Pdf.Text.TableAbsorber();
                absorber.Visit(page);
                foreach (AbsorbedTable table in absorber.TableList)
                {
                    int row_cnt = 0;                    
                    foreach (AbsorbedRow row in table.RowList)
                    {
                        int col_cnt = 0;
                        DataRow dr = dt_rst.NewRow();
                        foreach (AbsorbedCell cell in row.CellList)
                        {
                            TextFragment textfragment = new TextFragment();
                            TextFragmentCollection textFragmentCollection = cell.TextFragments;
                            string txt = "";
                            foreach (TextFragment fragment in textFragmentCollection)
                            {                                
                                foreach (TextSegment seg in fragment.Segments)
                                {
                                    txt += seg.Text;
                                }
                            }
                            if (row_cnt == 0) Global.AddColDt(dt_rst, txt, 0);
                            else
                            {
                                dr[col_cnt] = txt;
                            }
                            col_cnt++;
                        }
                        if (row_cnt != 0) dt_rst.Rows.Add(dr);                        
                        row_cnt++;
                    }                    
                }
            }
            
            //빈열 제거
            dt_rst = Global.DelEmptyColumn(dt_rst);            

            //결측행 제거
            dt_rst = Global.DelLittleRow(dt_rst, 1);

            //칼럼명을 뒤에서 바꾸기 위해 임시로 다른걸로 바꿔둠
            for (int i = 0; i < dt_rst.Columns.Count; i++)
                dt_rst.Columns[i].ColumnName = $"c{i.ToString()}";

            for (int i = 0; i < dt_rst.Columns.Count; i++)
            {
                if (i == 0) dt_rst.Columns[i].ColumnName = "Index";
                if (i == 1) dt_rst.Columns[i].ColumnName = "RT";
                if (i == 2) dt_rst.Columns[i].ColumnName = "Area";
                if (i == 3) dt_rst.Columns[i].ColumnName = "% Area";
                if (i == 4) dt_rst.Columns[i].ColumnName = "Height";
            }

            return dt_rst;
        }
    }

    public class MakeRawTable_SDC : ITable
    {
        public DataTable Extract_Table(string path)
        {
            PDDocument doc = PDDocument.load(path);
            PDFTextStripper stripper = new PDFTextStripper();
            string all_txt = stripper.getText(doc);
            DataTable dt_rst = new DataTable();
            int cnt = 0;
            foreach (string str in all_txt.Split(new string[] { "\r\n" }, StringSplitOptions.None))
            {
                if (ChkFrontThreeWord_IsNum(str))
                {
                    if (cnt == 0)
                    {
                        for (int i =0;i<str.Split(' ').Count(); i++)
                            Global.AddColDt(dt_rst, $"Column{i.ToString()}", 0);
                    }
                    
                    DataRow dr = dt_rst.NewRow();
                    int cnt2 = 0;
                    foreach (string s in str.Split(' '))
                    {
                        if (s != "") dr[cnt2] = s;
                        cnt2++;
                    }   

                    dt_rst.Rows.Add(dr);
                    cnt++;
                }
            }

            doc.close();

            //빈열 제거
            dt_rst = Global.DelEmptyColumn(dt_rst);

            //결측행 제거
            dt_rst = Global.DelLittleRow(dt_rst, 1);

            for (int i = 0; i < dt_rst.Columns.Count; i++)            
                dt_rst.Columns[i].ColumnName = $"c{i.ToString()}";


            //필요칼럼 추출
            dt_rst = dt_rst.DefaultView.ToTable(false, new string[] { "c0", "c1", "c3", "c5", "c4" });            

            for (int i = 0; i < dt_rst.Columns.Count; i++)
            {
                if (i == 0) dt_rst.Columns[i].ColumnName = "Index";
                if (i == 1) dt_rst.Columns[i].ColumnName = "RT";
                if (i == 2) dt_rst.Columns[i].ColumnName = "Area";
                if (i == 3) dt_rst.Columns[i].ColumnName = "% Area";
                if (i == 4) dt_rst.Columns[i].ColumnName = "Height";
            }            

            //return dt
            return dt_rst;
        }

        private bool ChkFrontThreeWord_IsNum(string str)
        {
            //앞의 세 단어가 모두 숫자이면 true, 데이터 행으로 반환
            int flag = 0;            
            double rst;
            string[] arr = str.Split(' ');

            foreach (string s in str.Split(' '))
            {
                if (s.Contains("Total")) break;               
                    
                if (Double.TryParse(s, out rst)) flag++;
                if (flag == 3) return true;
            }
            return false;
        }
    }

    public class MakeTableAll
    {
        private decimal De(string str)
        {
            return Decimal.Parse(str);
        }

        public DataTable GetRawTable(ITable itable, string path)
        {
            DataTable dt = itable.Extract_Table(path);

            //double형으로 변환
            dt = this.ConverDoubleTable(dt);

            //RRT 테이블 생성
            dt = this.MakeRRTColumn(dt);

            return dt;
        }

        public DataTable MakeImpurityTable(DataTable dt_raw, DataTable dt_ref, decimal limit)
        {
            DataTable dt_ref_raw = dt_ref.Copy();
            DataTable dt_ref_raw_AllRow = dt_ref_raw.Copy();
            while (dt_ref_raw.Rows.Count > 2) dt_ref_raw.Rows.RemoveAt(2);
            int? peak_idx = Global.VlookupDt_Int(dt_raw, Convert.ToDouble(dt_raw.AsEnumerable().Max(row => row["Area"])), "Area", "Index");
            
            DataTable dt_abs = dt_ref_raw.Clone();
            foreach (DataRow dr in dt_raw.Rows)
            {
                //Peak 인덱스는 제외
                if (dt_raw.Rows.IndexOf(dr) == peak_idx) continue;
                DataRow new_dr = dt_abs.NewRow();
                for (int i = 0;i<dt_ref_raw.Columns.Count;i++)
                {
                    if (i == 0) new_dr[i] = $"{(dt_raw.Rows.IndexOf(dr) + 1).ToString()}";
                    else new_dr[i] = Math.Abs(De(dr["RRT"].ToString()) - De(dt_ref_raw.Rows[1][i].ToString()));
                }
                dt_abs.Rows.Add(new_dr);
            }

            DataTable dt_absChk = dt_ref_raw.Clone();
            foreach (DataRow dr in dt_abs.Rows)
            {   
                DataRow new_dr = dt_absChk.NewRow();
                for (int i =0;i<dt_ref_raw.Columns.Count; i++)
                {
                    if (i == 0) new_dr[i] = dr[0].ToString();
                    else new_dr[i] = (De(dr[i].ToString()) <= limit) ? dr[i].ToString() : "";
                }
                dt_absChk.Rows.Add(new_dr);
            }            

            DataTable dt_imp = dt_ref_raw.Clone();
            DataRow dr_imp = dt_imp.NewRow();
            dt_imp.Rows.Add(dr_imp);            

            //불순물정하기
            foreach (DataRow dr in dt_absChk.Rows)
            {   
                int row_cnt = 0;
                int col_idx = 0;
                for (int i = 1;i<dt_absChk.Columns.Count;i++)
                {
                    if (dr[i].ToString() != "")
                    {
                        row_cnt++;
                        col_idx = i;
                    }
                }

                if (row_cnt == 0)
                {
                    dt_imp = AddDtNewImp(dt_imp, dr[0].ToString());
                                     
                }   
                else if (row_cnt == 1)
                {                   
                    //limit보다 낮은 abs값이 한 인덱스에 하나만 존재
                    Dictionary<string, decimal> dict = new Dictionary<string, decimal>();                    
                    for (int row = 0;row< dt_absChk.Rows.Count;row++)
                    {
                        if (dt_absChk.Rows[row][col_idx].ToString() != "")
                        {
                            dict.Add(dt_absChk.Rows[row][0].ToString(), De(dt_absChk.Rows[row][col_idx].ToString()));                            
                        }
                    }

                    if (dict.Count == 1)
                    {
                        //고유불순물
                        dr_imp[col_idx] = dr[0].ToString();
                    }
                    else
                    {
                        //한열에 두개
                        string minValueKey = dict.Aggregate((x, y) => x.Value < y.Value ? x : y).Key;
                        //중복불순물
                        if (dr[0].ToString() == minValueKey)
                        {
                            dr_imp[col_idx] = dr[0].ToString();
                        }
                        else
                        {
                            //신규
                            dt_imp = AddDtNewImp(dt_imp, dr[0].ToString());
                        }
                    }
                }
                else
                {
                    //한 행에 두개
                    //중복불순물
                    Dictionary<int, decimal> dict = new Dictionary<int, decimal>();
                    for (int col = 1; col < dt_absChk.Columns.Count; col++)
                    {
                        if (dr[col].ToString() != "")
                        {
                            dict.Add(col, De(dr[col].ToString()));
                        }
                    }

                    int minValueKey = dict.Aggregate((x, y) => x.Value < y.Value ? x : y).Key;
                    dr_imp[minValueKey] = dr[0].ToString();
                }
            }

            dt_imp.Columns.RemoveAt(0);

            //Peak
            Global.AddColDt_Index(dt_imp, "Peak", 0);
            dr_imp["Peak"] = $"{peak_idx}";
            
            //RRT 등 넣기
            Global.AddColDt_Index(dt_imp, "Item", 0);            
            dt_imp = AddItemRow(dt_imp, dt_raw, "% Area", 0);            
            dt_imp = AddItemRow(dt_imp, dt_raw, "RRT", 0);            
            dt_imp = AddItemRow(dt_imp, dt_raw, "RT", 0);

            //dt_ref            
            dt_imp = AddItemRefRow(dt_imp, dt_ref_raw_AllRow, "RRT", 0);            
            dt_imp = AddItemRefRow(dt_imp, dt_ref_raw_AllRow, "RT", 0);            

            for (int i = 0; i< dt_imp.Rows.Count; i++)
            {
                if (i == 0) dt_imp.Rows[i][0] = "Ref RT";
                if (i == 1) dt_imp.Rows[i][0] = "Ref RRT";
                if (i == 2) dt_imp.Rows[i][0] = "QC RT";
                if (i == 3) dt_imp.Rows[i][0] = "QC RRT";
                if (i == 4) dt_imp.Rows[i][0] = "QC %Area";
                if (i == 5) dt_imp.Rows[i][0] = "Index";
            }

            return dt_imp;
        }

        private DataTable AddItemRefRow(DataTable dt_imp, DataTable dt_ref, string item, int pos)
        {
            DataRow new_dr = dt_imp.NewRow();
            new_dr[0] = item;            
            for (int i =1;i<dt_imp.Columns.Count; i++)
            {
                if (dt_imp.Columns[i].ColumnName == "Peak")
                {
                    if (item == "RT") new_dr["Peak"] = dt_ref.Rows[dt_ref.Rows.Count - 1][1].ToString();
                    if (item == "RRT") new_dr["Peak"] = 1;
                }
                else
                {
                    string col_name = dt_imp.Columns[i].ColumnName;
                    double? val = Global.HVlookupDt(dt_ref, item, col_name);
                    new_dr[i] = val;
                }
                
            }
            dt_imp.Rows.InsertAt(new_dr, pos);
            return dt_imp;
        }

        private DataTable AddItemRow(DataTable dt_imp, DataTable dt_raw, string item, int pos)
        {
            DataRow new_dr = dt_imp.NewRow();
            new_dr[0] = item;
            for (int i = 1; i < dt_imp.Columns.Count; i++)
            {   
                if (dt_imp.Rows[dt_imp.Rows.Count - 1][i].ToString() != "")
                {
                    double idx = Double.Parse(dt_imp.Rows[dt_imp.Rows.Count - 1][i].ToString());
                    double? val = Global.VlookupDt(dt_raw, idx, "Index", item);
                    new_dr[i] = Math.Round(double.Parse(val.ToString()), 3);
                }
                else new_dr[i] = "";                
            }
            dt_imp.Rows.InsertAt(new_dr, pos);
            return dt_imp;
        }

        private DataTable AddDtNewImp(DataTable dt_imp, string Index)
        {
            //신규불순물
            if (!dt_imp.Columns[dt_imp.Columns.Count - 1].ToString().Contains("New")) Global.AddColDt(dt_imp, "New1", 0);
            else
            {
                int new_cnt = Int32.Parse(dt_imp.Columns[dt_imp.Columns.Count - 1].ToString().Replace("New","")) + 1;
                Global.AddColDt(dt_imp, $"New{new_cnt}", 0);
            }
            dt_imp.Rows[0][dt_imp.Columns.Count - 1] = Index;

            return dt_imp;
        }

        //private
        private DataTable MakeRRTColumn(DataTable dt)
        {
            DataColumn dataColumn = new DataColumn(columnName: "RRT", dataType : typeof(double));
            dt.Columns.Add("RRT", typeof(double)).SetOrdinal(2);            
            double? max_rt = Global.VlookupDt(dt, Convert.ToDouble(dt.AsEnumerable().Max(row => row["Area"])), "Area", "RT");
            foreach (DataRow dr in dt.Rows)
            {
                dr["RRT"] = double.Parse(dr["RT"].ToString()) / max_rt;
            }

            return dt;
        }

        private DataTable ConverDoubleTable(DataTable dt)
        {
            DataTable ret = dt.Clone();
            foreach (DataColumn col in ret.Columns)            
                col.DataType = typeof(double);

            foreach (DataRow dr in dt.Rows)
            {
                DataRow new_dr = ret.NewRow();
                for (int i = 0; i < ret.Columns.Count; i++)
                {
                    new_dr[i] = double.Parse(dr[i].ToString());
                }
                ret.Rows.Add(new_dr);
            }

            return ret;
        }
    }
}