   M S B A M L   `   `   `  ���� B  >LGchem2, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:LGchem2LGchem2      LGchem2.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     'localclr-namespace:LGchem2       � Title$  LGChem2��$
��1550��$	��900�� � WindowStartupLocation$ CenterScreen=� � UseLayoutRounding .  � 
ResizeMode$ CanMinimize=� � 	Resources %   LGchem2.ChkeckToHeaderConverter   ChkeckToHeaderConverter��%   &	          .��� $��10 10 10 10q�.+��� $��Stretch=�$��Stretch=�.@��� $��메인��.���� $��0 5 0 0q� �� Orientation$ 
Horizontal=�.+��� ���� $��메인��.���� .+��� $ 
Horizontal=�.+��� -    ��Name$ btn_select_pdf��$��PDF 파일 추가��$��5 0 0 0q�$	��100��$��25��
�������$
��Left=� .  J� Effect 3  -System.Windows.Media.Effects.DropShadowEffect     ShadowDepth$ 2Q�	   	Direction$		 315Q�
   Color$
 Black��   Opacity$	 0.5Q��� -   $ btn_run��$��시작��$��20 0 0 0q�$	��100��$��25��
��� � �$
��Left=� .    $ 2Q�$		 315Q�$
 Black��$	 0.5Q��� -   $ 
btn_result��$��결과 폴더��$��20 0 0 0q�$	��100��$��25��
�������$
��Left=� .    $ 2Q�$		 315Q�$
 Black��$	 0.5Q��� -   $ 	btn_reset��$��	초기화��$��20 0 0 0q�$	��100��$��25��
�������$
��Left=� .    $ 2Q�$		 315Q�$
 Black��$	 0.5Q��� -   $ btn_result_open��$��결과 파일 열기��$��20 0 0 0q�$	��100��$��25��
�������$
��Left=� .    $ 2Q�$		 315Q�$
 Black��$	 0.5Q��� -    ��Name$ tb_result_path��$	��310��$��Center=�$��14(�$��10 0 0 0q��� $ 
Horizontal=�.+�� $��5 10 0 0q�.+�� -    �Name$ pgb_run��$	��870��$��25�� � Minimum$ 0Q� � Maximum$	 100Q�$
��Left=� � Value �� *val+ �� UpdateSourceTrigger$ PropertyChanged=� � IsIndeterminate �� *isindertate+$ PropertyChanged=��� -    ��Name$ pgb_text��$��Center=�$��5 0 0 0q�$��Center=����� *str+$ PropertyChanged=�   LGchem2.perBusySpinner  -	     Name$ Spin_Control��$��20��
������ 
����  �$��Right=�$��Center=�$��5 11 0 0q��� $ 
Horizontal=�.+��� -
   $��lb_time��$��소요시간��$	��100��$��5 6 0 0q��� -    ��Name$ chk_pdf_ole��$6��0PDF 삽입 여부(미체크시 속도 빨라짐)��$��	30 10 0 0q� d� 	IsChecked$
 TrueY��� $ 
Horizontal=�.+�� $��Stretch=�.+��� -   $��list_pdf��$��5 5 10 0q� �� SelectionMode$ Extended=�$	��900��$	��635��$��14(� J� 	AllowDrop . ����  �� View �� .P��� $��	파일명�� �� Width$	 300�� �� DisplayMemberBinding �� *
pdf_name+�� $��경로�� �� *
pdf_path+�� $ 
Horizontal=�$��0 10 0 0q�.+��� $��RRT 한계값���� -   $ tb_rrt_limit��$	��100��$��0.015��$��Center=�$��14(�$��10 0 0 0q��� $��작업 폴더��$��10 0 0 0q��� -   $ tb_workFolder��$	��400��$��Center=�$��14(�$��10 0 0 0q��� -   $ btn_workFolderSelect��$��작업폴더 선택��$��20 0 0 0q�$	��100��$��25��
�������$
��Left=� .    $ 2Q�$		 315Q�$
 Black��$	 0.5Q��� -   $ btn_workFolderOpen��$��작업폴더 열기��$��20 0 0 0q�$	��100��$��25��
�������$
��Left=� .    $ 2Q�$		 315Q�$
 Black��$	 0.5Q��� ���� $��1개 결과 확인��.���� .+��� $ 
Horizontal=�.+��� $��Raw 데이터��$��5 0 0 0 q��� -   $ chk_raw_header��$��칼럼명 복사 여부��$��5 6 0 0q�$
 TrueY�&   System.Windows.Controls.DataGrid  -   $��dg_raw_result��$��5 5 5 0q�$��14(�$	��550��$	��350��$��Stretch=�   SelectionUnit$
 Cell=�   
IsReadOnly . ����    ClipboardCopyMode ��  �� ElementName$ chk_raw_header��  �� Path$  	IsChecked�! �� 	Converter#! [  �� $ 
Horizontal=�.+��� $��불순물 데이터��$��5 5 0 0q��� -   $ chk_imp_header��$��칼럼명 복사 여부��$��5 11 0 0q�$
 TrueY�  -   $��dg_imp_result��$��5 5 5 0q�$��14(�$	��550��$	��350��$��Stretch=�$
 Cell=� . ����  �� $ chk_imp_header��$  	IsChecked�#! [  �� $��	사용법��.���� -   $ tb_tib��$��Stretch=�$��Stretch=�$��5 5 5 5q�" �� 
IsReadOnly" . $��20(�