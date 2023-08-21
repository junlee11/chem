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
using static LGchem2.ExcelControl;

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
        private Dictionary<string, double?> dic_ref_spc = new Dictionary<string, double?>();
        private Dictionary<string, double?> dic_ref_lcl = new Dictionary<string, double?>();
        private Pgb_Val pgb_val = new Pgb_Val();
        string cur_path = System.IO.Directory.GetCurrentDirectory();
        List<Model_pdf> model_Pdfs = new List<Model_pdf>();
        private TestEnum testEnum = new TestEnum();

        public enum TestEnum
        {
            run,
            test
        }

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

            testEnum = TestEnum.test;

            this.tb_tib.Text = "<초기 폴더 및 레퍼런스 설정>\n" +
                "1. 작업 폴더 설정 : 작업폴더 선택 버튼을 클릭하여 Ref파일과 결과 파일이 저장될 폴더를 선택한다.\n" +
                "2. 작업 폴더 선택을 하면 작업 폴더 경로가 설정되고 작업폴더 열기로 작업폴더를 들어갈 수 있다.\n" +
                "3. 작업 폴더의 Ref폴더에 Ref.xlsx 파일을 레퍼런스로 넣어야 한다.\n\n" +
                "<PDF 선택>\n" +
                "1. PDF 파일 추가 또는 파일을 드래그하여 PDF 파일을 추가할 수 있다. 다른 확장자의 파일은 추가되지 않는다.\n" +
                "2. 삭제할때는 삭제할 파일을 선택하여 키보드 Delete를 누르면 된다. 모두 삭제하려면 Ctrl+A로 모두 선택한 뒤 Delete를 누른다.\n" +
                "3. 결과 파일은 작업폴더의 Result폴더에 남는다. 결과폴더를 통해 폴더를 들어가거나 결과 파일 열기로 결과 파일을 열 수 있다.\n\n" +
                "<레퍼런스 파일 설명>\n" +
                "1. 레퍼런스 시트명이 PDF 파일명 @ 앞 문자열에 포함이 되면 레퍼런스가 있다고 인식한다.\n" +
                "2. 레퍼런스 표를 작성할때 반드시 1열은 RT , RRT , SPEC , LCL , REF_RT 순서로 되어 있어야 한다. (대소문자 구별함)\n" +
                "3. SPEC, LCL, REF_RT는 공란이어도 상관없다.\n" +
                "4. SpecOut은 빨간색, LCL Out은 파란색으로 Peak의 % Area 글씨가 표시된다.\n" +
                "5. 레퍼런스 표는 반드시 1행 1열부터 시작해야 한다.\n\n" +
                "<데이터 관련 설명>\n" +
                "1. 1개의 파일만 실행하면 오른쪽 표에 Raw데이터, 불순물 표가 표시된다. (1개만 빠르게 확인할 때 사용하라는 의도)\n" +
                "2. PDF 파일이 레퍼런스에 포함되는 파일이면 3개의 표가 생성된다.\n" +
                "→ Raw데이터 표, 불순물 환산 표, RefRRT - RawRRT 값을 계산하여 RRT limit값 이내인 값만 표시한 표(불순물 환산 표를 만드는데 사용됨)\n" +
                "3. Ref가 없는 PDF라면 불순물 환산 표는 생성되지 않는다." +
                "4. SPC 탭에는 Ref를 기준으로 여러 파일의 % Area 데이터가 정리된다.";

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
                    //MessageBox.Show("존재하는 폴더로 다시 선택하세요");
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
            if (this.tb_workFolder.Text == "")
            {
                MessageBox.Show("작업폴더가 설정되지 않았습니다.");
                return;
            }

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
            if (e.Key == Key.Delete)
            {
                if (this.list_pdf.SelectedItems.Count == 0) return;

                foreach (var item in this.list_pdf.SelectedItems)
                {
                    var cast_item = (Model_pdf)item;
                    model_Pdfs.Remove(cast_item);
                }
                this.list_pdf.Items.Refresh();
            }
        }

        private void btn_run_Click(object sender, RoutedEventArgs e)
        {
            //인터락
            if (this.tb_workFolder.Text == "")
            {
                MessageBox.Show("작업 폴더가 설정되지 않았습니다.");
                return;
            }

            if (!File.Exists($"{this.tb_workFolder.Text}\\LGchem2\\Ref\\Ref.xlsx"))
            {
                MessageBox.Show($"{this.tb_workFolder.Text}\\LGchem2\\Ref\\Ref.xlsx 파일이 없습니다.");
                return;
            }
            
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
            PgbControl(0, "작업 시작");

            List<string> list_path = new List<string>();
            foreach (var item in this.list_pdf.Items) list_path.Add(((Model_pdf)item).pdf_path);

            Dictionary<string, object> dic_src = new Dictionary<string, object>();
            dic_src.Add("limit", limit);
            dic_src.Add("list_path", list_path);
            dic_src.Add("chk_pdfole", (bool)this.chk_pdf_ole.IsChecked);
                        
            this.dg_raw_result.DataContext = null;
            this.dg_imp_result.DataContext = null;
            this.Spin_Control.Visibility = Visibility.Visible;
            this.lb_time.Content = "소요시간";
            this.btn_run.IsEnabled = false;

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
            try
            {
                Dictionary<string, object> dic_src = (Dictionary<string, object>)o;
                decimal limit = (decimal)dic_src["limit"];
                List<string> list_path = (List<string>)dic_src["list_path"];
                List<PdfDt> list_pdfdt = new List<PdfDt>();
                bool chk_pdfole = (bool)dic_src["chk_pdfole"];

                ExcelControl excelControl = new ExcelControl();

                double pgb_all_val = list_path.Count * 4;

                Stopwatch sw = new Stopwatch();
                sw.Start();

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

                    if (dt_raw == null) throw new Exception("Raw PDF 데이터가 표로 변환되지 않습니다.");

                    PgbControl(((list_path.IndexOf(path) + 1) / pgb_all_val) * 80, $"{fileName} raw 데이터테이블 생성완료");

                    //2. 레퍼런스 체크 : 엑셀 시트명이 파일명의 골뱅이 앞에 있는 문자열에 속해야함
                    //3. rrt 테이블 생성

                    DataTable dt_imp = new DataTable();

                    //dic_ref에 재료에 해당하는 레퍼런스 dt 있는지 검사해서 최초이면 레퍼런스 적재
                    string key = "";

                    DataTable dt_val = new DataTable();
                    (key, dt_val) = excelControl.GetDic_SheetContentTable(refExcel_path, Path.GetFileName(path).Split('@')[0]);
                    if (key != "" && !Global.ChkStrInDicKey(fileName.Split('@')[0], dic_ref))
                        dic_ref.Add(key, dt_val);


                    DataTable dt_ref = null;
                    double? spc = null;
                    double? lcl = null;
                    DataTable dt_absChk = null;

                    //그 재료가 레퍼런스에 속하는지 검사해서 속하면 임퓨리티 테이블 생성
                    if (Global.ChkStrInDicKey(fileName.Split('@')[0], dic_ref))
                    {
                        //레퍼런스에 있을때 불순물 테이블 만듬
                        dt_ref = Global.GetdtInDicKey(fileName.Split('@')[0], dic_ref);
                        double temp;
                        //spc 
                        if (!dic_ref_spc.ContainsKey(key))
                        {
                            if (dt_ref.Rows[2][1].ToString() == "") spc = null;
                            else if (Double.TryParse(dt_ref.Rows[2][1].ToString(), out temp)) spc = temp;
                            else spc = null;
                            dic_ref_spc.Add(key, spc);
                        }

                        //lcl
                        if (!dic_ref_lcl.ContainsKey(key))
                        {
                            if (dt_ref.Rows[3][1].ToString() == "") lcl = null;
                            else if (Double.TryParse(dt_ref.Rows[3][1].ToString(), out temp)) lcl = temp;
                            else lcl = null;
                            dic_ref_lcl.Add(key, lcl);
                        }

                        (dt_imp, dt_absChk) = makeTableAll.MakeImpurityTable(dt_raw, dt_ref, limit);
                    }
                    else dt_imp = null;

                    if (!dic_ref_spc.ContainsKey("")) dic_ref_spc.Add("", null);
                    if (!dic_ref_lcl.ContainsKey("")) dic_ref_lcl.Add("", null);

                    PgbControl(((list_path.IndexOf(path) + 2) / pgb_all_val) * 80, $"{fileName} imp 데이터테이블 생성완료");
                    //list에 담기
                    list_pdfdt.Add(new PdfDt()
                    {
                        pdf_path = path,
                        pdf_name = fileName,
                        dt_raw = dt_raw,
                        dt_imp = dt_imp,
                        dt_ref = dt_ref,
                        spc = dic_ref_spc[key],
                        lcl = dic_ref_lcl[key],
                        ref_name = key,
                        dt_absChk = dt_absChk
                    });
                }

                int row_Idx = 2;
                int col_rawIdx = 2;
                int col_impIdx = 10;

                DataTable dt_spc = new DataTable();
                Dictionary<string, DataTable> dic_spc = new Dictionary<string, DataTable>();

                //spc용 dic 생성
                foreach (KeyValuePair<string, DataTable> items in dic_ref) if (!dic_spc.ContainsKey(items.Key)) dic_spc.Add(items.Key, null);
                foreach (PdfDt item in list_pdfdt)
                {
                    //Spc 테이블 뼈대 만들기                
                    if (dic_spc.ContainsKey(item.ref_name))
                    {
                        if (dic_spc[item.ref_name] == null || dic_spc[item.ref_name].Columns.Count < item.dt_imp.Columns.Count)
                        {
                            dic_spc[item.ref_name] = item.dt_imp.Clone();
                            dic_spc[item.ref_name].ImportRow(item.dt_imp.Rows[0]);      //Ref RT                        
                            dic_spc[item.ref_name].ImportRow(item.dt_imp.Rows[1]);      //Ref RRT                        
                        }
                    }
                }
                //sw.Stop();
                //MessageBox.Show($"Dt 생성 {sw.ElapsedMilliseconds}ms");

                //Stopwatch sw2 = new Stopwatch();
                //sw2.Start();

                //아래부터 엑셀
                Excel.Application application = null;
                Excel.Workbook workBook = null;
                string excelFileName = (list_pdfdt.Count == 0) ? "" : excelFileName = $"{result_path}\\result_{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.xlsx";

                try
                {
                    if (excelFileName == "") throw new Exception("저장할 fileName이 없습니다.");
                    //Excel 프로그램 실행
                    application = new Excel.Application();
                    //Excel 화면 띄우기 옵션
                    application.Visible = false;
                    //파일로부터 불러오기
                    workBook = application.Workbooks.Add();

                    if (list_pdfdt.Count != 0)
                    {
                        workBook.Worksheets[1].Name = "Main";
                        workBook.Worksheets.Add(After: workBook.Sheets[workBook.Sheets.Count]);
                        workBook.Worksheets[2].Name = "SPC";
                    }

                    foreach (PdfDt item in list_pdfdt)
                    {
                        //4. raw테이블과 rrt 테이블 배치하기
                        //5. pdf 첨부                
                        ExcelControl.Spec spc = new ExcelControl.Spec();
                        if (item.dt_ref == null) spc = ExcelControl.Spec.nospc;
                        else
                        {
                            double peak_val = double.Parse(item.dt_imp.Rows[4]["Peak"].ToString());

                            if (item.lcl == null || item.spc == null) spc = ExcelControl.Spec.nospc;
                            else if (item.lcl >= item.spc)
                            {
                                if (peak_val >= item.lcl) spc = ExcelControl.Spec.spcIn;
                                else if (peak_val >= item.spc) spc = ExcelControl.Spec.lclOut;
                                else spc = ExcelControl.Spec.spcOut;
                            }
                            else spc = ExcelControl.Spec.nospc;
                        }

                        excelControl.DataTableToExcelQuick(item.dt_raw, workBook, "Main", row_Idx, col_rawIdx, spc, item.pdf_path);
                        //excelControl.DataTableToExcel(item.dt_raw, excelFileName, "Main", row_Idx, col_rawIdx, spc, item.pdf_path);
                        PgbControl(((list_path.Count * 2 + list_pdfdt.IndexOf(item) + 1) / pgb_all_val) * 80, $"{item.pdf_name} raw 테이블 엑셀 생성완료");

                        excelControl.DataTableToExcelQuick(item.dt_imp, workBook, "Main", row_Idx, col_impIdx, spc);
                        //excelControl.DataTableToExcel(item.dt_imp, excelFileName, "Main", row_Idx, col_impIdx, spc);
                        PgbControl(((list_path.Count * 2 + list_pdfdt.IndexOf(item) + 2) / pgb_all_val) * 80, $"{item.pdf_name} imp 테이블 엑셀 생성완료");

                        if (testEnum == TestEnum.test && item.dt_imp != null)
                        {
                            //dt_absChk 테스트용
                            excelControl.DataTableToExcelQuick(item.dt_absChk, workBook, "Main", row_Idx, col_impIdx + item.dt_imp.Columns.Count + 2, Spec.nospc);
                            //excelControl.DataTableToExcel(item.dt_absChk, excelFileName, "Main", row_Idx, col_impIdx + item.dt_imp.Columns.Count + 2, Spec.nospc);
                        }

                        if (item.dt_imp != null)
                        {
                            if (item.dt_raw.Rows.Count >= item.dt_imp.Rows.Count) row_Idx += item.dt_raw.Rows.Count + 3;
                            else row_Idx += item.dt_imp.Rows.Count + 3;
                        }
                        else row_Idx += item.dt_raw.Rows.Count + 3;

                        //Spc 테이블 만들기
                        if (dic_spc.ContainsKey(item.ref_name))
                        {
                            DataRow dr = item.dt_imp.Rows[4];
                            dic_spc[item.ref_name].ImportRow(dr);
                            dic_spc[item.ref_name].Rows[dic_spc[item.ref_name].Rows.Count - 1][0] = item.pdf_name.Replace(".pdf", "");
                        }
                    }

                    //6. SPC 적재
                    row_Idx = 1;
                    col_impIdx = 1;
                    foreach (KeyValuePair<string, DataTable> items in dic_spc)
                    {
                        if (items.Value == null) continue;
                        excelControl.DataTableToExcelQuick(items.Value, workBook, "SPC", row_Idx, col_rawIdx, Spec.nospc);
                        //excelControl.DataTableToExcel(items.Value, excelFileName, "SPC", row_Idx, col_rawIdx, Spec.nospc);
                        row_Idx = row_Idx + items.Value.Rows.Count + 2;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    if (excelFileName != "")
                    {
                        workBook.SaveAs(excelFileName);
                        workBook.Close();
                        application.Quit();        // 엑셀 어플리케이션 종료
                                                   //오브젝트 해제
                        Global.ReleaseExcelObject(workBook);
                        Global.ReleaseExcelObject(application);
                    }
                }

                PgbControl(90, $"PDF 개체 삽입중");
                if (chk_pdfole)
                {
                    //PDF 개체 삽입
                    Excel.Application app_macro = null;
                    Excel.Workbook wb_macro = null;

                    try
                    {
                        row_Idx = 2;
                        foreach (PdfDt item in list_pdfdt)
                        {
                            //엑셀 매크로
                            app_macro = new Excel.Application();
                            app_macro.Visible = false;
                            wb_macro = app_macro.Workbooks.Open($"{cur_path}\\macro.xlsm");
                            Excel.Worksheet ws_macro = wb_macro.ActiveSheet;

                            //PDF 객체 매크로로 삽입
                            ws_macro.Cells[1, 2] = item.pdf_path;
                            ws_macro.Cells[3, 2] = item.pdf_name;
                            ws_macro.Cells[4, 2] = excelFileName;
                            ws_macro.Cells[5, 2] = row_Idx;
                            ws_macro.Cells[6, 2] = 1;
                            //Call VBA code
                            app_macro.Run("load_ole");

                            if (item.dt_imp != null)
                            {
                                if (item.dt_raw.Rows.Count >= item.dt_imp.Rows.Count) row_Idx += item.dt_raw.Rows.Count + 3;
                                else row_Idx += item.dt_imp.Rows.Count + 3;
                            }
                            else row_Idx += item.dt_raw.Rows.Count + 3;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    finally
                    {
                        if (excelFileName != "")
                        {
                            wb_macro.Close(false);
                            app_macro.Quit();        // 엑셀 어플리케이션 종료
                                                     //오브젝트 해제
                            Global.ReleaseExcelObject(wb_macro);
                            Global.ReleaseExcelObject(app_macro);
                        }
                    }
                }

                if (list_pdfdt.Count == 1)
                {
                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                    {
                        this.dg_raw_result.DataContext = list_pdfdt[0].dt_raw;
                        this.dg_imp_result.DataContext = list_pdfdt[0].dt_imp;
                    }));
                }
                //sw2.Stop();
                //MessageBox.Show($"엑셀 생성 {sw2.ElapsedMilliseconds}ms");

                PgbControl(100, $"전체 완료");
                if (excelFileName != "")
                {
                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                    {
                        this.tb_result_path.Text = excelFileName;
                    }));
                }

                sw.Stop();
                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                {
                    this.Spin_Control.Visibility = Visibility.Hidden;
                    this.lb_time.Content = $"{sw.ElapsedMilliseconds}ms 소요";
                    this.btn_run.IsEnabled = true;
                    MessageBox.Show("완료되었습니다.");
                    this.tb_result_path.Focus();
                    this.tb_result_path.Select(this.tb_result_path.Text.Length, 0);
                }));
            }
            catch(Exception ex)
            {
                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                {
                    PgbControl(0, "에러발생");
                    this.Spin_Control.Visibility = Visibility.Hidden;
                    this.lb_time.Content = $"에러발생";
                    this.btn_run.IsEnabled = true;
                    this.tb_result_path.Text = "";                    
                                        
                }));
                MessageBox.Show(ex.ToString());
            }
            
        }

        private void btn_result_Click(object sender, RoutedEventArgs e)
        {
            try { Process.Start(result_path); }
            catch { }
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
                if (Path.GetExtension(file).ToUpper() == ".PDF") model_Pdfs.Add(new Model_pdf() { pdf_name = Path.GetFileName(file), pdf_path = file });
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

        private void btn_reset_Click(object sender, RoutedEventArgs e)
        {
            this.tb_result_path.Text = "";
            this.list_pdf.DataContext = null;
            PgbControl(0, "작업 대기");
            this.dg_imp_result.DataContext = null;
            this.dg_raw_result.DataContext = null;
        }
    }
}