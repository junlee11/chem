using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Path = System.IO.Path;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection.Emit;

namespace LGchem2
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private string basic_path = @"C:\LGchem2";
        private string result_path = @"C:\LGchem2\Result";
        private string ref_path = @"C:\LGchem2\Ref";
        private string refExcel_path = @"C:\LGchem2\Ref\Ref.xlsx";
        private Dictionary<string, DataTable> dic_ref = new Dictionary<string, DataTable>();
        private Pgb_Val pgb_val = new Pgb_Val();

        public MainWindow()
        {
            InitializeComponent();
            DirectoryInfo dic = new DirectoryInfo(basic_path);
            if (!dic.Exists) dic.Create();

            dic = new DirectoryInfo(result_path);
            if (!dic.Exists) dic.Create();

            dic = new DirectoryInfo(ref_path);
            if (!dic.Exists) dic.Create();

            this.pgb_run.DataContext = pgb_val;
            this.pgb_text.DataContext = pgb_val;

            //this.Loaded += new RoutedEventHandler(MainPage_Loaded);

            ////pdf test
            //string path = @"C:\Users\USER\Desktop\코드\lgchem_sample\230809\LG-008(W)\LG-008@008P23002@C[99.802%].pdf";

            //MakeTable_LGD makeTable_LGD = new MakeTable_LGD();
            //MakeTable_SDC makeTable_SDC = new MakeTable_SDC();
            //MakeTableAll makeTableAll = new MakeTableAll();

            //DataTable dt_raw = new DataTable();
            //try
            //{
            //    dt_raw = makeTableAll.GetRawTable(makeTable_LGD, path);                
            //}
            //catch (System.IndexOutOfRangeException)
            //{
            //    dt_raw = makeTableAll.GetRawTable(makeTable_SDC, path);
            //}
            ////double형으로 변환
            //dt_raw = makeTableAll.ConverDoubleTable(dt_raw);

            ////RRT 테이블 생성
            //dt_raw = makeTableAll.MakeRRTColumn(dt_raw);
        }

        private void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void btn_select_pdf_Click(object sender, RoutedEventArgs e)
        {
            //PDF 선택
            OpenFileDialog ofdlg = new OpenFileDialog();
            {
                ofdlg.InitialDirectory = @"C:\driver";   // 기본 폴더
                ofdlg.Filter = // 필터설정
                    "PDF Files | *.pdf; *.PDF;" + "| " +
                    "All Files |*.*";

                ofdlg.CheckFileExists = true;   // 파일 존재여부확인
                ofdlg.CheckPathExists = true;   // 폴더 존재여부확인
                ofdlg.Multiselect = true;

                // 파일 열기 (값의 유무 확인)
                if (ofdlg.ShowDialog().GetValueOrDefault())
                {
                    List<string> list_pdf = ofdlg.FileNames.ToList();

                    List<Model_pdf> model_Pdfs = new List<Model_pdf>();
                    foreach (string file in list_pdf)
                    {
                        model_Pdfs.Add(new Model_pdf() { pdf_name = Path.GetFileName(file), pdf_path = file });
                    }
                    this.list_pdf.DataContext = model_Pdfs;
                }
            }
        }

        private void btn_run_Click(object sender, RoutedEventArgs e)
        {
            //인터락
            if (this.list_pdf.Items.Count == 0)
            {
                MessageBox.Show("PDF 파일을 선택하세요");
                return;
            }
                
            decimal limit;
            if (!Decimal.TryParse(this.tb_rrt_limit.Text, out limit))
            {
                MessageBox.Show("RRT 한계값은 실수가 입력되어야 합니다.");
                return;
            }

            List<string> list_path = new List<string>();
            foreach (var item in this.list_pdf.Items) list_path.Add(((Model_pdf)item).pdf_path);

            Dictionary<string, object> dic_src = new Dictionary<string, object>();
            dic_src.Add("limit", limit);
            dic_src.Add("list_path", list_path);

            this.pgb_run.IsIndeterminate = true;
            
            //PDF 만들기
            Thread th = new Thread(new ParameterizedThreadStart(th_pdf));
            th.IsBackground = true;
            th.Start(dic_src);
        }

        private void th_pdf(object o)
        {
            Dictionary<string, object>dic_src = (Dictionary<string, object>)o;
            decimal limit = (decimal)dic_src["limit"];
            List<string> list_path = (List<string>)dic_src["list_path"];
            List<PdfDt> list_pdfdt = new List<PdfDt>();

            ExcelControl excelControl = new ExcelControl();

            //pdf 반복
            foreach (string path in list_path)
            {
                string fileName = Path.GetFileName(path);
                Debug.WriteLine("good");

                //1. pdf 변환
                MakeRawTable_LGD makeTable_LGD = new MakeRawTable_LGD();
                MakeRawTable_SDC makeTable_SDC = new MakeRawTable_SDC();
                MakeTableAll makeTableAll = new MakeTableAll();

                DataTable dt_raw = new DataTable();
                try
                {
                    dt_raw = makeTableAll.GetRawTable(makeTable_LGD, path);
                }
                catch (System.IndexOutOfRangeException)
                {
                    dt_raw = makeTableAll.GetRawTable(makeTable_SDC, path);
                }

                //2. 레퍼런스 체크 : 엑셀 시트명이 파일명의 골뱅이 앞에 있는 문자열에 속해야함
                //3. rrt 테이블 생성
                
                DataTable dt_imp = new DataTable();

                //dic_ref에 재료에 해당하는 레퍼런스 dt 있는지 검사해서 최초이면 레퍼런스 적재
                if (!Global.ChkStrInDicKey(fileName.Split('@')[0], dic_ref))
                {
                    string key = "";
                    DataTable dt_val = new DataTable();
                    (key, dt_val) = excelControl.GetDic_SheetContentTable(refExcel_path, Path.GetFileName(path).Split('@')[0]);
                    if (key != "") dic_ref.Add(key, dt_val);
                }

                DataTable dt_ref = null;
                //그 재료가 레퍼런스에 속하는지 검사해서 속하면 임퓨리티 테이블 생성
                if (Global.ChkStrInDicKey(fileName.Split('@')[0], dic_ref))
                {
                    //레퍼런스에 있을때 불순물 테이블 만듬
                    dt_ref = Global.GetdtInDicKey(fileName.Split('@')[0], dic_ref);
                    dt_imp = makeTableAll.MakeImpurityTable(dt_raw, dt_ref, limit);
                }
                else
                {
                    dt_imp = null;
                }

                //list에 담기
                list_pdfdt.Add(new PdfDt() { pdf_path = path, pdf_name = fileName, dt_raw = dt_raw, dt_imp = dt_imp, dt_ref = dt_ref });
            }
            
            string excelFileName = "";
            if (list_pdfdt.Count != 0)
            {
                excelFileName = $"{result_path}\\result_{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.xlsx";
                excelControl.SaveExcelFile(excelFileName);
            }

            int row_Idx = 2;
            int col_rawIdx = 2;
            int col_impIdx = 10;
            
            foreach (PdfDt item in list_pdfdt)
            {
                //4. raw테이블과 rrt 테이블 배치하기
                //5. pdf 첨부
                ExcelControl.Spec spc = new ExcelControl.Spec();
                Global.print_DataTable(item.dt_ref);
                if (item.dt_ref == null) spc = ExcelControl.Spec.nospc;
                else
                {
                    if (item.dt_ref.Rows[2][0].ToString() == "" || item.dt_ref.Rows[3][0].ToString() == "")
                    {
                        spc = ExcelControl.Spec.nospc;
                    }
                    double peak_val = double.Parse(item.dt_imp.Rows[4]["Peak"].ToString());
                    double lcl;
                    double spec;

                    if (!Double.TryParse(item.dt_ref.Rows[2][1].ToString(), out spec) || Double.TryParse(item.dt_ref.Rows[3][1].ToString(), out lcl))
                    {
                        spc = ExcelControl.Spec.nospc;
                    }
                    else
                    {
                        spec = Double.Parse(item.dt_ref.Rows[2][1].ToString());
                        lcl = Double.Parse(item.dt_ref.Rows[3][1].ToString());

                        if (peak_val >= lcl) spc = ExcelControl.Spec.spcIn;
                        else if (peak_val >= spec) spc = ExcelControl.Spec.lclOut;
                        else spc = ExcelControl.Spec.spcOut;
                    }
                }
                
                Debug.WriteLine(item.pdf_name);
                Debug.WriteLine(item.pdf_path);
                Global.print_DataTable(item.dt_raw);
                Global.print_DataTable(item.dt_imp);

                excelControl.DataTableToExcel(item.dt_raw, excelFileName, row_Idx, col_rawIdx, spc, item.pdf_path);
                excelControl.DataTableToExcel(item.dt_imp, excelFileName, row_Idx, col_impIdx, spc);

                if (item.dt_raw.Rows.Count >= item.dt_imp.Rows.Count) row_Idx += item.dt_raw.Rows.Count + 7;
                else row_Idx += item.dt_imp.Rows.Count + 5;
            }
            Debug.WriteLine("good");

            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                this.pgb_run.IsIndeterminate = false;
            }));
        }

        private void btn_result_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(result_path);
        }
    }
}