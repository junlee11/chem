using Microsoft.Win32;
using System;
using System.Collections.Generic;
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

        public MainWindow()
        {
            InitializeComponent();
            DirectoryInfo dic = new DirectoryInfo(basic_path);
            if (!dic.Exists) dic.Create();

            dic = new DirectoryInfo(result_path);
            if (!dic.Exists) dic.Create();

            dic = new DirectoryInfo(ref_path);
            if (!dic.Exists) dic.Create();
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
            if (this.list_pdf.SelectedItems.Count == 0) return;
            double rrt_limit;
            if (!Double.TryParse(this.tb_rrt_limit.Text, out rrt_limit))
            {
                MessageBox.Show("RRT 한계값은 실수가 입력되어야 합니다.");
                return;
            }

            List<string> list_path = new List<string>();
            foreach (var item in this.list_pdf.SelectedItems) list_path.Add(((Model_pdf)item).pdf_path);

            Dictionary<string, object> dic_src = new Dictionary<string, object>();
            dic_src.Add("rrt_limit", rrt_limit);
            dic_src.Add("list_path", list_path);

            this.pgb_run.IsIndeterminate = true;
            
            //PDF 만들기
            Thread th = new Thread(new ParameterizedThreadStart(th_pdf));
            th.IsBackground = true;
            th.Start();
        }

        private void th_pdf(object o)
        {
            Dictionary<string, object>dic_src = (Dictionary<string, object>)o;
            double rrt_limit = (double)dic_src["rrt_limit"];
            List<string> list_path = (List<string>)dic_src["list_path"];
            
            foreach (string path in list_path)
            {
                
            }

            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
            {
                this.pgb_run.IsIndeterminate = false;

            }));
        }
    }
}