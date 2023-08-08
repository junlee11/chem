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
                    foreach (AbsorbedRow row in table.RowList)
                    {
                        foreach (AbsorbedCell cell in row.CellList)
                        {
                            TextFragment textfragment = new TextFragment();
                            TextFragmentCollection textFragmentCollection = cell.TextFragments;
                            foreach (TextFragment fragment in textFragmentCollection)
                            {
                                string txt = "";
                                foreach (TextSegment seg in fragment.Segments)
                                {
                                    txt += seg.Text;
                                }
                                Debug.WriteLine(txt);
                            }
                        }
                    }
                }
            }

            return dt_rst;
        }
    }
}
