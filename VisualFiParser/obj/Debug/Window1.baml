   M S B A M L   `   `   `  ����I  EVisualFiParser, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null/clr-namespace:VisualFiParserVisualFiParser  5   	       VisualFiParser.Window1   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   .localclr-namespace:VisualFiParser     5   	     � Title$  Visual FiParser - Info��5   	   $	��526��6(   $	��678��65    � WindowStartupLocation$ CenterScreen=�5	   	    � 
ResizeMode$ CanMinimize=�5
   	   .��5      � .+�5   
   �� -   5         ��Name$ 	textBlock��5      $��
20,35,22,0q�6'   $
��Wrap=�6;   $	��Top=�6O   $��86��6g   �� 6t   $�Visual FiParser��6x   �� 6�   $� consente di compilare ��6�   �� 6�   $�rapidamente ��6�   �� 6�   $#�i moduli delle gare di Apnea ��6�   �� 6�   $�
sveltendo ��6�   �� 6  $�una pratica ��6  �� 6-  $�molto ripetitiva��61  �� 6K  $�.��6O  �� 6Z  �� 6f  �� 6l  �� 6x  $M�GVisual FiParser è una applicazione Open Source sviluppata in C# e .Net��6|  �� 6�  $"�, rilasciata con licenza MIT��6�  �� 6�  $
� ed ��6�  �� 6  -   5         ��Name$ git_link��5     $-��'https://github.com/ozeta/VisualFiParser?�6<  .�6s  �� $�ospitata su GitHub��6w  6�  6�  �� 5   
   -   5         ��Name$ image��5      $
��Left=�6   $��225,121,0,0q�6:   $	��100��6O   $	��100��6[   $	��Top=�6h   �� 5   
   -   5        $ GOPR0461_1_jpg��5      $��
20,151,0,0q�6(   $��GOPR0461_1.jpg��6<   $
��Fill=�6T   $
��Left=�6c   $	��115��6~   $	��115��6�   $	��Top=�6�   �� 5   
   -   5        $ textBlock_Copy��5      $��150,151,22,0q�6,   $
��Wrap=�6B   $��85��6V   $	��Top=�6b   �� 6{   $N�HMi chiamo Marco Carrozzo, in questo periodo sto terminando gli studi in ��6   �� 6�   $�I��6�   �� 6�   $:�4nformatica, pratico apnea dal 2007 e sono diventato ��6�   �� 6"  $�A��6&  �� 61  $V�Pllenatore Indoor/Outdoor nel 2013. Ho una lunga esperienza agonistica nel nuoto ��65  �� 6�  $)�#e sono in categoria Elité dal 2013��6�  �� 6�  $� ad oggi��6�  �� 6�  $�.��6�  6�  �� 5   
   -   5         ��Name$ label��5      $��Chi sono��6   $
��Left=�62   $��
10,120,0,0q�6M   $	��Top=�6a   $
��Bold"�6y   �� 5   
   -   5        $ app_name��5      $��Fiparser��6"   $
��Left=�65   $��	10,10,0,0q�6P   $	��Top=�6c   $
��Bold"�6{   �� 5   
   -   5        $ txtbFooterRight��5      $��Bottom=�6-   $	��108��6H   $��452,0,10,22q�6U   �� 5      -	   5        $ pp_link��5       �� TextDecorations$
 None{�6)   $)��#https://www.paypal.me/MarcoCarrozzo?�6@   .�6�   �� $� ��6�   �� 6�   .F�5      �� -
   5        $ 
paypal_jpg��5   A   $	��179��6   $
��Fill=�6   $��
paypal.jpg��6-   $��92��6U   5      6    6,   �� 5   
   -   5        $ textBlock_Copy1��5      $��20,0,42,158q�6-   $
��Wrap=�6B   $��59��6V   $��Bottom=�6b   �� 6~   $T�NHo sviluppato FiParser come esigenza personale, per sbrigare più rapidamente ��6�   �� 6�   $�( in un click ) ��6�   �� 6�   $0�*le pratiche burocratiche legate alle gare.��6�   �� 6/  $N�H Ho rilasciato pubblicamente questa applicazione come Esercizio di stile��63  �� 6�  $� e per aiutare ��6�  �� 6�  $�qualche collega ��6�  �� 6�  $
�di..��6�  �� 6�  $�.��6�  �� 6�  $�burocrazia :-)��6�  �� 6�  �� 6	  �� 6  6  �� 5   
   -   5        $ textBlock_Copy2��5      $��	20,0,0,98q�6-   $
��Wrap=�6@   $��37��6T   $��Bottom=�6`   $
��Left=�6{   $	��420��6�   $l��fSe hai apprezzato il mio lavoro, puoi fare una donazione per supportare lo sviluppo dell'applicazione.��6�   �� 5   
   -   5        $ textBlock_Copy3��5      $��20,0,0,3q�6-   $
��Wrap=�6?   $��80��6S   $��Bottom=�6_   $
��Left=�6z   $	��303��6�   �� 6�   -   5        $ fb_link��5   �   $$��https://www.facebook.com/ozeta?�6�   .�6  �� $ �Marco Carrozzo su Facebook��6  6.  �� 69  $� ��6=  �� 6G  �� 6S  �� 6Y  �� 6e  -   5        $ 	mail_link��5   �  $#��mailto:mar.carrozzo@gmail.com?�6o  .�6�  �� $�mar.carrozzo@gmail.com��6�  6�  �� 6�  �� 6�  $� ��6  6  5      5      