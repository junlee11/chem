   M S B A M L   `   `   `  ����B  >LGchem2, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null!clr-namespace:LGchem2LGchem2  5          LGchem2.perBusySpinner   5      \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      &ctrlclr-namespace:LGchem2     5      $	��120��5      ��5      ��   �� RelativeSource  �� *Self+ �� Path$ Height� �� Mode$ TwoWay=�.��5	      (� $��Stretch=�5
   	   $��Stretch=�5   	   .��5   
   � ��6   �� $ 
Background�  ��  �� Mode$ FindAncestor=� �� AncestorType   LGchem2.perBusySpinner"  .+�5      �� $��Center=�5      $��Center=�5      
������  J� RenderTransformOrigin$ 0.5,0.5&�5       � UseLayoutRounding .  ��5      �� $ Diameter�  �� $ FindAncestor=�"  ��5      �� $ Diameter�  �� $ FindAncestor=�"   � 	Resources 5       �� 
TargetType	 �� Property
 �� Value��%�  'A�       �� 5      " A�6   .�5      �� 	 � ��
 61   �� $ 
Foreground�  �� $ FindAncestor=�"  �� 5      	 � ��
 63   �� $ ItemDiameter�  �� $ FindAncestor=�"  �� 5      	 �@��Stretch$

 Fill=�64   �� 5      	 � ��$
 0��6<   �� 5      	 � ��
 62   �� $ ItemDiameter�  �� $ FindAncestor=�"  5      5      .+�5       A� -   5         A�Name$ Item1��5!       J� Opacity$	 1.0Q�5"      A� 5#      -   5        $ Item2��5$      $
 0.92Q�5%      A� 5&      -   5        $ Item3��5'      $
 0.84Q�5(      A� 5)      -   5        $ Item4��5*      $
 0.76Q�5+      A� 5,      -   5        $ Item5��5-      $
 0.68Q�5.      A� 5/      -   5        $ Item6��50      $
 0.60Q�51      A� 52      -   5        $ Item7��53      $
 0.52Q�54      A� 55      -   5        $ Item8��56      $
 0.44Q�57      A� 58      -	   5        $ Item9��59      $
 0.36Q�5:      A� 5;      -
   5        $ Item10��5<      $
 0.28Q�5=      A� 5>      -   5        $ Item11��5?      $
 0.20Q�5@      A� 5A      -   5        $ Item12��5B      $
 0.12Q�5C      z�5E      �� 5F      -   5        SpinnerRotateTransform�� �� Angle$ 0Q�5F   F   5H      5I      5J      5L      