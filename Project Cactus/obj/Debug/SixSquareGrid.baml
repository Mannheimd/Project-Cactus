   M S B A M L   `   `   `  ����B  >Cactus, Version=1.0.12.0, Culture=neutral, PublicKeyToken=nullclr-namespace:CactusCactus  5   	       Cactus.SixSquareGrid   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Cactus     5   	     � Title$  Six Square Grid��5   	   $	��430��6!   $	��680��6.   .��5	      � .+�5
   
   �� .��5      � $��5,5,5,5q�6   .+�5      � $��5,5,5,5q�6   $	��180��6$   $	��200��61   ���#223a5e.+�5      �� $
��Auto��6   
�������$��Permissions��68   $
��Bold"�6N   $��20(�6`   $��Center=�6n   $	��Top=�6�   �� 5      $
��Wrap=�6   $��5,35,5,0q�60   $	��Top=�6B    �� TextAlignment$ Center=�6Z   �� 5      
�������.�6-   ?=Can the problem be recreated when logging on to Windows as an6l    �� 5      
������� �� TextDecorations$ 	Underline{�6-   $
��Bold"�6I   .�6[   Administrator6j    �� 5      
�������.�6-   or61    �� 5      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   different user?6l   5      �� 5      $
��Wrap=�6   $��5,105,105,0q�60   $	��Top=�6E   $ Center=�6]   �� 5      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Yes6`   �� 5      �� 5      
�������.�6-   
Probably67    �� 5      
�������$ 	Underline{�6-   .�6I   not6N    �� 5      
�������.�6-   a User specific problem6F   5      �� 5      $
��Wrap=�6   $��105,105,5,0q�60   $	��Top=�6E   $ Center=�6]   �� 5      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   No6_   �� 5      �� 5      
�������.�6-   
Probably67    �� 5      
�������$ 	Underline{�6-   .�6I   is6M    �� 5       
�������.�6-   a User specific problem6F   5!      5"      � 5#      $��5,5,5,5q�6   $	��180��6$   $	��200��61   ���#223a5e.+�5$      �� $
��Auto��6   
�������$��Computer��68   $
��Bold"�6K   $��20(�6]   $��Center=�6k   $	��Top=�6�   �� 5%      $
��Wrap=�6   $��5,35,5,0q�60   $	��Top=�6B   $ Center=�6Z   �� 5&      
�������.�6-   !Can the problem be recreated on6N    �� 5'      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   another computer?6n   �� 5(      �� 5)      
�������.�6-   (If using a networked setup)6K   5*      �� 5+      $
��Wrap=�6   $��5,105,105,0q�60   $	��Top=�6E   $ Center=�6]   �� 5,      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Yes6`   �� 5-      �� 5.      
�������.�6-   
Probably67    �� 5/      
�������$ 	Underline{�6-   .�6I   not6N    �� 50      
�������.�6-   a Computer specific problem6J   51      �� 52      $
��Wrap=�6   $��105,105,5,0q�60   $	��Top=�6E   $ Center=�6]   �� 53      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   No6_   �� 54      �� 55      
�������.�6-   
Probably67    �� 56      
�������$ 	Underline{�6-   .�6I   is6M    �� 57      
�������.�6-   a Computer specific problem6J   58      59      � 5:      $��5,5,5,5q�6   $	��180��6$   $	��200��61   ���#223a5e.+�5;      �� $
��Auto��6   
�������$��AV/Firewall��68   $
��Bold"�6N   $��20(�6`   $��Center=�6n   $	��Top=�6�   �� 5<      $
��Wrap=�6   $��5,35,5,0q�60   $	��Top=�6B   $ Center=�6Z   �� 5=      
�������.�6-   #!Can the problem be recreated with6P    �� 5>      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Antivirus and Firewall6s    �� 5?      
�������.�6-   :8software disabled/uninstalled, or with exceptions added?6g   5@      �� 5A      $
��Wrap=�6   $��5,105,105,0q�60   $	��Top=�6E   $ Center=�6]   �� 5B      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Yes6`   �� 5C      �� 5D      
�������.�6-   
Probably67    �� 5E      
�������$ 	Underline{�6-   .�6I   not6N    �� 5F      
�������.�6-   !an AV/Firewall specific problem6N   5G      �� 5H      $
��Wrap=�6   $��105,105,5,0q�60   $	��Top=�6E   $ Center=�6]   �� 5I      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   No6_   �� 5J      �� 5K      
�������.�6-   
Probably67    �� 5L      
�������$ 	Underline{�6-   .�6I   is6M    �� 5M      
�������.�6-   !an AV/Firewall specific problem6N   5N      5O      � 5P      $��5,5,5,5q�6   $	��180��6$   $	��200��61   ���#223a5e.+�5Q      �� $
��Auto��6   
�������$��Application��68   $
��Bold"�6N   $��20(�6`   $��Center=�6n   $	��Top=�6�   �� 5R      $
��Wrap=�6   $��5,35,5,0q�60   $	��Top=�6B   $ Center=�6Z   �� 5S      
�������.�6-   &$Can the problem be recreated on your6S    �� 5T      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   own computer6i    �� 5U      
�������.�6-   )'using the same configuration and build?6V   5V      �� 5W      $
��Wrap=�6   $��5,105,105,0q�60   $	��Top=�6E   $ Center=�6]   �� 5X      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Yes6`   �� 5Y      �� 5Z      
�������.�6-   
Probably67    �� 5[      
�������$ 	Underline{�6-   .�6I   is6M    �� 5\      
�������.�6-   caused by the application6H   5]      �� 5^      $
��Wrap=�6   $��105,105,5,0q�60   $	��Top=�6E   $ Center=�6]   �� 5_      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   No6_   �� 5`      �� 5a      
�������.�6-   
Probably67    �� 5b      
�������$ 	Underline{�6-   .�6I   not6N    �� 5c      
�������.�6-   caused by the application6H   5d      5e      � 5f      $��5,5,5,5q�6   $	��180��6$   $	��200��61   ���#223a5e.+�5g      �� $
��Auto��6   
�������$��Database��68   $
��Bold"�6K   $��20(�6]   $��Center=�6k   $	��Top=�6�   �� 5h      $
��Wrap=�6   $��5,35,5,0q�60   $	��Top=�6B   $ Center=�6Z   �� 5i      
�������.�6-   !Can the problem be recreated on6N    �� 5j      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   	another6d    �� 5k      
�������.�6-   or61    �� 5l      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   a new6b    �� 5m      
�������.�6-   	database?68   5n      �� 5o      $
��Wrap=�6   $��5,105,105,0q�60   $	��Top=�6E   $ Center=�6]   �� 5p      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Yes6`   �� 5q      �� 5r      
�������.�6-   
Probably67    �� 5s      
�������$ 	Underline{�6-   .�6I   not6N    �� 5t      
�������.�6-   a database specific problem6J   5u      �� 5v      $
��Wrap=�6   $��105,105,5,0q�60   $	��Top=�6E   $ Center=�6]   �� 5w      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   No6_   �� 5x      �� 5y      
�������.�6-   
Probably67    �� 5z      
�������$ 	Underline{�6-   .�6I   is6M    �� 5{      
�������.�6-   a database specific problem6J   5|      5}      � 5~      $��5,5,5,5q�6   $	��180��6$   $	��200��61   ���#223a5e.+�5      �� $
��Auto��6   
�������$��Database User��68   $
��Bold"�6P   $��20(�6b   $��Center=�6p   $	��Top=�6�   �� 5�      $
��Wrap=�6   $��5,35,5,0q�60   $	��Top=�6B   $ Center=�6Z   �� 5�      
�������.�6-   =;Can the problem be recreated when logging in as a different6j    �� 5�      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   database user?6k   5�      �� 5�      $
��Wrap=�6   $��5,105,105,0q�60   $	��Top=�6E   $ Center=�6]   �� 5�      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   Yes6`   �� 5�      �� 5�      
�������.�6-   
Probably67    �� 5�      
�������$ 	Underline{�6-   .�6I   not6N    �� 5�      
�������.�6-   specific to the database user6L   5�      �� 5�      $
��Wrap=�6   $��105,105,5,0q�60   $	��Top=�6E   $ Center=�6]   �� 5�      
�������$ 	Underline{�6-   $
��Bold"�6I   .�6[   No6_   �� 5�      �� 5�      
�������.�6-   
Probably67    �� 5�      
�������$ 	Underline{�6-   .�6I   is6M    �� 5�      
�������.�6-   specific to the database user6L   5�      5�      5�      5�      5�      5�      