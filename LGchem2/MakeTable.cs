using Aspose.Pdf.Text;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text; 
using System.Threading.Tasks;

namespace LGchem2
{
    public class MakeTable
    {
        public DataTable Extract_Table(string path)
        {
            DataTable dt_rst = new DataTable();
            // Load source PDF document
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(path);
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
                            if (row_cnt == 0) Global.AddColDt(dt_rst, txt);
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
            
            return dt_rst;
        }
    }
}
