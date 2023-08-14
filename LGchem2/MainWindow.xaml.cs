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
using java.security.cert;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace LGchem2
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private string basic_path;
        private string result_path;
        private string ref_path;
        private string refExcel_path;
        private Dictionary<string, DataTable> dic_ref = new Dictionary<string, DataTable>();
        private Pgb_Val pgb_val = new Pgb_Val();
        string cur_path = System.IO.Directory.GetCurrentDirectory();
        List<Model_pdf> model_Pdfs = new List<Model_pdf>();

        public MainWindow()
        {
            InitializeComponent();

            this.Spin_Control.Visibility = Visibility.Hidden;            

            FileInfo fileInfo = new FileInfo(cur_path + "\\tar_path.txt");
            if (!fileInfo.Exists)
            {
                File.WriteAllText(fileInfo.FullName, "", Encoding.UTF8);
            }
            this.tb_workFolder.Text = File.ReadAllText(fileInfo.FullName);

            WorkFolderValidate();

            this.pgb_run.DataContext = pgb_val;
            this.pgb_text.DataContext = pgb_val;

            //this.Loaded += new RoutedEventHandler(MainPage_Loaded);            
        }

        private void WorkFolderValidate()
        {
            string work_folder = this.tb_workFolder.Text;
            try
            {
                DirectoryInfo dic = new DirectoryInfo(work_folder);
                if (dic.Exists)
                {
                    basic_path = $"{work_folder}\\LGchem2";
                    result_path = $"{work_folder}\\LGchem2\\Result";
                    ref_path = $"{work_folder}\\LGchem2\\Ref";
                    refExcel_path = $"{work_folder}\\LGchem2\\Ref\\Ref.xlsx";

                    dic = new DirectoryInfo(basic_path);
                    if (!dic.Exists) dic.Create();

                    dic = new DirectoryInfo(result_path);
                    if (!dic.Exists) dic.Create();

                    dic = new DirectoryInfo(ref_path);
                    if (!dic.Exists) dic.Create();
                }
                else
                {
                    MessageBox.Show("존재하는 폴더로 다시 선택하세요");
                }
            }
            catch (System.ArgumentException)
            {
                //pass
            }
            
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
                    
                    foreach (string file in list_pdf)
                    {
                        model_Pdfs.Add(new Model_pdf() { pdf_name = Path.GetFileName(file), pdf_path = file });
                    }
                    this.list_pdf.DataContext = model_Pdfs;
                    this.list_pdf.Items.Refresh();
                }
            }
        }

        private void list_pdf_KeyDown(object sender, KeyEventArgs e)
        {
            if (this.list_pdf.SelectedItems.Count == 0) return;

            foreach (var item in this.list_pdf.SelectedItems)
            {
                var cast_item = (Model_pdf)item;
                model_Pdfs.Remove(cast_item);                
            }
            this.list_pdf.Items.Refresh();
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
            this.tb_result_path.Text = "";

            List<string> list_path = new List<string>();
            foreach (var item in this.list_pdf.Items) list_path.Add(((Model_pdf)item).pdf_path);

            Dictionary<string, object> dic_src = new Dictionary<string, object>();
            dic_src.Add("limit", limit);
            dic_src.Add("list_path", list_path);
                        
            this.dg_raw_result.DataContext = null;
            this.dg_imp_result.DataContext = null;
            this.Spin_Control.Visibility = Visibility.Visible;

            //PDF 만들기
            Thread th = new Thread(new ParameterizedThreadStart(th_pdf));
            th.IsBackground = true;
            th.Start(dic_src);
        }

        private void PgbControl(double val, string str)
        {
            pgb_val.val = val;
            pgb_val.str = str;
        }

        private void th_pdf(object o)
        {
            Dictionary<string, object>dic_src = (Dictionary<string, object>)o;
            decimal limit = (decimal)dic_src["limit"];
            List<string> list_path = (List<string>)dic_src["list_path"];
            List<PdfDt> list_pdfdt = new List<PdfDt>();

            ExcelControl excelControl = new ExcelControl();

            double pgb_all_val = list_path.Count * 4;
            double pgb_sep_val = 0;

            //pdf 반복
            foreach (string path in list_path)
            {
                string fileName = Path.GetFileName(path);                

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

                PgbControl(((list_path.IndexOf(path) + 1) / pgb_all_val) * 100, $"{fileName} raw 데이터테이블 생성완료");                

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
                double? spc = null;
                double? lcl = null;
                
                //그 재료가 레퍼런스에 속하는지 검사해서 속하면 임퓨리티 테이블 생성
                if (Global.ChkStrInDicKey(fileName.Split('@')[0], dic_ref))
                {   
                    //레퍼런스에 있을때 불순물 테이블 만듬
                    dt_ref = Global.GetdtInDicKey(fileName.Split('@')[0], dic_ref);
                    double temp;
                    if (dt_ref.Rows[2][1].ToString() == "") spc = null;
                    else if (Double.TryParse(dt_ref.Rows[2][1].ToString(), out temp)) spc = temp;
                    else spc = null;                    

                    if (dt_ref.Rows[3][1].ToString() == "") lcl = null;
                    else if (Double.TryParse(dt_ref.Rows[3][1].ToString(), out temp)) lcl = temp;
                    else lcl = null;

                    dt_imp = makeTableAll.MakeImpurityTable(dt_raw, dt_ref, limit);                    
                }
                else
                {
                    dt_imp = null;
                }

                PgbControl(((list_path.IndexOf(path) + 2) / pgb_all_val) * 100, $"{fileName} imp 데이터테이블 생성완료");
                //list에 담기
                list_pdfdt.Add(new PdfDt() { pdf_path = path, pdf_name = fileName, dt_raw = dt_raw, dt_imp = dt_imp, dt_ref = dt_ref, spc = spc, lcl = lcl });
            }
            
            string excelFileName = "";
            if (list_pdfdt.Count != 0)
            {
                excelFileName = $"{result_path}\\result_{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.xlsx";
                excelControl.SaveExcelFile(excelFileName);
                excelControl.SheetNameChange(excelFileName, 1, "Main");
                excelControl.SheetNameChange(excelFileName, 2, "SPC");
            }

            int row_Idx = 2;
            int col_rawIdx = 2;
            int col_impIdx = 10;

            DataTable dt_spc = new DataTable();            
            
            foreach (PdfDt item in list_pdfdt)
            {
                //4. raw테이블과 rrt 테이블 배치하기
                //5. pdf 첨부
                ExcelControl.Spec spc = new ExcelControl.Spec();                                
                if (item.dt_ref == null) spc = ExcelControl.Spec.nospc;
                else
                {                    
                    double peak_val = double.Parse(item.dt_imp.Rows[4]["Peak"].ToString());                    

                    if (item.lcl == null || item.spc == null)
                    {
                        spc = ExcelControl.Spec.nospc;
                    }
                    else if (item.lcl >= item.spc)
                    {
                        if (peak_val >= item.lcl) spc = ExcelControl.Spec.spcIn;
                        else if (peak_val >= item.spc) spc = ExcelControl.Spec.lclOut;
                        else spc = ExcelControl.Spec.spcOut;
                    }
                    else
                    {
                        spc = ExcelControl.Spec.nospc;
                    }
                }

                excelControl.DataTableToExcel(item.dt_raw, excelFileName, "Main", row_Idx, col_rawIdx, spc, item.pdf_path);
                PgbControl(((list_path.Count * 2 + list_pdfdt.IndexOf(item) + 1) / pgb_all_val) * 100, $"{item.pdf_name} raw 엑셀 생성완료");

                excelControl.DataTableToExcel(item.dt_imp, excelFileName, "Main", row_Idx, col_impIdx, spc);
                PgbControl(((list_path.Count * 2 + list_pdfdt.IndexOf(item) + 2) / pgb_all_val) * 100, $"{item.pdf_name} raw 엑셀 생성완료");

                if (item.dt_raw.Rows.Count >= item.dt_imp.Rows.Count) row_Idx += item.dt_raw.Rows.Count + 3;
                else row_Idx += item.dt_imp.Rows.Count + 3;
            }            

            if (list_pdfdt.Count == 1)
            {
                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                {
                    this.dg_raw_result.DataContext = list_pdfdt[0].dt_raw;
                    this.dg_imp_result.DataContext = list_pdfdt[0].dt_imp;
                }));                
            }

            PgbControl(100, $"전체 완료");

            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {                
                this.Spin_Control.Visibility = Visibility.Hidden;                
                MessageBox.Show("완료되었습니다.");                
            }));

            if (excelFileName != "")
            {
                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                {
                    this.tb_result_path.Text = excelFileName;
                }));
            }
        }

        private void btn_result_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(result_path);
        }

        private void btn_workFolderSelect_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    this.tb_workFolder.Text = dialog.FileName;

                    FileInfo fileInfo = new FileInfo(cur_path + "\\tar_path.txt");                    
                    File.WriteAllText(fileInfo.FullName, dialog.FileName, Encoding.UTF8);
                    WorkFolderValidate();
                }
            }
        }

        private void btn_workFolderOpen_Click(object sender, RoutedEventArgs e)
        {
            try {  Process.Start(basic_path); }
            catch { }            
        }

        private void btn_result_open_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start(this.tb_result_path.Text); }
            catch { }
        }

        private void list_pdf_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string file in files)
            {                
                if (Path.GetExtension(file) != ".pdf") model_Pdfs.Add(new Model_pdf() { pdf_name = Path.GetFileName(file), pdf_path = file });
            }
            this.list_pdf.DataContext = model_Pdfs;
            this.list_pdf.Items.Refresh();
        }

        private void list_pdf_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (this.list_pdf.SelectedItems.Count != 1) return;

            Model_pdf model_Pdf = (Model_pdf)this.list_pdf.SelectedItem;
            try { Process.Start(model_Pdf.pdf_path); }
            catch { }
        }
    }
}