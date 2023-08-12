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

            Debug.WriteLine(pdfDocument.Pages.Count);
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
            Global.print_DataTable(dt_raw);
            Global.print_DataTable(dt_ref);

            DataTable dt_rst = new DataTable();

            return dt_rst;
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