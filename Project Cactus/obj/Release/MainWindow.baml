   M S B A M L   `   `   `  ���� B  >Cactus, Version=1.0.12.0, Culture=neutral, PublicKeyToken=nullclr-namespace:CactusCactus      Cactus.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     &localclr-namespace:Cactus     -     � Title$  Cactus��$	��585��$	��387��$	��290��$	��250�� J� CommandBindings  � 	Resources  �� MergedDictionaries �� Source��  �� $ WatermarkResource.xaml?�%    .���  � RowDefinitions �� $��20 ��� $��55 ��� $��43 ��� $��* ��� $��43 �.+�]� $��0��.+�o� $	��Top=���Z� .@�m� $��_File��.@�m� -   $��_Exit��m� $��_Edit��.@�m�  m� Command$	 Cut��m� $
 Copy��m� $ Paste��m� $��_Tools��.@�m� -   $��_Helpdesk Checklist��m� -   $��_Six Square Grid��m� $��_Help��.@�m� -   $��How To��m� -   $��View notes backup��m� -   $��About���� $��1��$��10,5,0,0q�$!��swiftpage_logo_50pxhigh.png��$
��Left=�$
��None=�� -    �Name$ ticketLoggingButtonsTop��$��2��$	��Top=�$��43��.+��� -	    ��Name$ startTimer_Button��$��Start Timer��$��	10,10,0,0q�$
��Left=�$	��Top=�$��76��	 � ToolTip$$	 Starts the call duration timer���� -
   $ resetForm_Button��$��
Reset Form��$��	0,10,10,0q�$��Right=�$	��Top=�$��76��$7	 1Resets the form to default - you cannot undo this���� $��Bottom=�$��Stretch=�$��1�����	#FFFE5000�� $��3��$
��Auto=�$��0,0,0,0q�$
��auto��.��� -   $ ticketLoggingFields�� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+�� -   $ reasonForCall_Grid��$��1��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -   
 ��Name$
 reasonForCall_Label��$��Reason for call:��$<	 6Provide a short summary of why the customer is calling��$
��Left=�$��	10,0,10,0q�$	��Top=��� -    ��Name$ reasonForCall_TextBox��$<	 6Provide a short summary of why the customer is calling��   WatermarkedTextBox#��[  
 � Tag$" Why is the customer calling?��$��stretch=�$
��Auto��$��
10,26,10,7q� �� TextWrapping$
 Wrap=� �� AcceptsReturn . � -   $ productInfoGrid��$��2��$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+�� -   $ product_Row��$��0��.+��� -   $
 product_Label��$��Product:��$
��Left=�$��10,0,0,0q�$	��Top=�$��26��$:	 4Select the product the customer is calling regarding���� -    ��Name$ product_ComboBox��$��Stretch=�$��	69,2,10,0q�$	��Top=�$
��auto��$��22��$:	 4Select the product the customer is calling regarding��� -   $ productName_Row��$��1��$y�	Collapsed=�.+��� -   $
 productName_Label��$��Name:��$
��Left=�$��10,0,0,0q�$	��Top=�$��26��$E	 ?Enter the name of the product the customer is calling regarding���� -   $ productName_TextBox��$��Stretch=�$��	69,2,10,0q�$	��Top=�$
��auto��$
��Auto�� . $E	 ?Enter the name of the product the customer is calling regarding��#��[  $% What product is the call about?��� -   $ versionText_Row��$��2��$y�	Collapsed=�.+��� -   $
 versionText_Label��$��Version:��$
��Left=�$��10,0,0,0q�$	��Top=�$��26��$	 Enter the product version���� -   $ versionText_TextBox��$��Stretch=�$��	69,2,10,0q�$	��Top=�$
��auto��$
��auto��$	 Enter the product version��#��[  $" What version is the product?��� -   $ productVersion_Row��$��3��$y�	Collapsed=�.+��� -   $
 productVersion_Label��$��Version:��$
��Left=�$��10,0,0,0q�$	��Top=�$��26��$ 	 Select the product version���� -   $ productVersion_ComboBox��$��Stretch=�$��	69,2,10,0q�$	��Top=�$
��auto��$��22��$ 	 Select the product version��� -   $ productUpdate_Row��$��4��$y�	Collapsed=�.+��� -   $
 productUpdate_Label��$��Update:��$
��Left=�$��40,0,0,0q�$	��Top=�$��26��$6	 0Enter the Update or HotFix number of the product���� -   $ productUpdate_TextBox��$��stretch=�$
��Auto��$��	99,2,10,0q�$
 Wrap=�$	��Top=� . $6	 0Enter the Update or HotFix number of the product��#��[  $ Update or Hotfix number��� -   $ os_Row��$��5��$y�	Collapsed=�.+��� -    $
 os_Label��$	��OS:��$��10,0,0,0q�$
��Left=�$	��Top=�$,	 &Select the customer's Operating System���� -!   $ os_ComboBox��$��	69,2,10,0q�$��Stretch=�$	��Top=�$,	 &Select the customer's Operating System��� -"   $ browser_Row��$��6��$y�	Collapsed=�.+��� -#   $
 browser_Label��$��Browser:��$��10,0,0,0q�$
��Left=�$	��Top=�$'	 !Select the customer's Web browser���� -$   $ browser_ComboBox��$��	69,2,10,0q�$��Stretch=�$	��Top=�$'	 !Select the customer's Web browser��� -%   $ browserVersion_Row��$��7��$y�	Collapsed=�.+��� -&   $
 browserVersion_Label��$��Version:��$��40,0,0,0q�$
��Left=�$	��Top=�$,	 &Enter the web browser's version number���� -'   $ browserVersion_TextBox��$��stretch=�$
��Auto��$��	99,2,10,0q�$
 Wrap=�$	��Top=� . $,	 &Enter the web browser's version number��#��[  $ Browser version number��� -(   $ accountName_Row��$��8��$y�	Collapsed=�.+��� -)   $
 accountName_Label��$��Account:��$��10,0,0,0q�$
��Left=�$	��Top=�$2	 ,Enter the customer's EMarketing account name���� -*   $ accountName_TextBox��$��stretch=�$
��Auto��$��	69,2,10,0q�$
 Wrap=�$	��Top=� . $2	 ,Enter the customer's EMarketing account name��#��[  $ EMarketing Account Name��� -+   $ sql_Row��$��9��$y�	Collapsed=�.+��� -,   $
 	sql_Label��$
��SQL:��$��10,0,0,0q�$
��Left=�$	��Top=�$.	 (Select the customer's SQL Server Version���� --   $ sql_ComboBox��$��	69,2,10,0q�$��Stretch=�$	��Top=�$.	 (Select the customer's SQL Server Version��� -.   $ 
office_Row��$��10��$y�	Collapsed=� �� $
��auto ��� $
��auto �.+�� $��0��.+��� -/   $
 office_Label��$��Office:��$��10,0,0,0q�$
��Left=�$	��Top=�$=	 7Select the version of Office that the customer is using���� -0   $ office_ComboBox��$��	69,2,10,0q�$��Stretch=�$	��Top=�$=	 7Select the version of Office that the customer is using��� -1   $ officeArchitecture_Row��$��1��$y�	Collapsed=�.+�	� -2    	�Name$* $officeArchitecture_RadioButton_32bit�� 	� 	GroupName$ officeArchitecture��$��32-bit��$��40,4,0,4q�$
��Left=�	� -3   $* $officeArchitecture_RadioButton_64bit��$ officeArchitecture��$��64-bit��$��	100,4,0,4q�$
��Left=�� -4   $ 	other_Row��$��11��$y�	Collapsed=�.+��� -5   $
 other_Label��$��Other:��$��10,0,0,0q�$
��Left=�$	��Top=�$�	 �Enter any other version information, such as for any software that the customer is using in conjunction with Swiftpage products. Antivirus and Firewall information should be entered here too.���� -6   $ other_TextBox��$��stretch=�$
��Auto��$��	69,2,10,0q�$
 Wrap=�$	��Top=� . $
��Auto=�$�	 �Enter any other version information, such as for any software that the customer is using in conjunction with Swiftpage products. Antivirus and Firewall information should be entered here too.��#��[  $/ )Enter any other environmental information��� -7   $ errorMessages_Grid��$��4��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -8   $
 errorMessages_Label��$��Error Messages:��$
��Left=�$��10,5,0,0q�$	��Top=�$7	 1Enter any error or alert messages exactly as seen���� -9   $ errorMessages_TextBox��$��stretch=�$
��Auto��$��
10,31,10,5q�$
 Wrap=�$	��Top=� . $7	 1Enter any error or alert messages exactly as seen��#��[  $# Exact error or alert messages��� -:   $  additionalInformation_Grid��$��5��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -;   $!
 additionalInformation_Label��$)��#Additional issue/query information:��$
��Left=�$��10,0,0,0q�$	��Top=�$G	 AEnter any additional information relating to the problem or query���� -<   $# additionalInformation_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $G	 AEnter any additional information relating to the problem or query��#��[  $$ Additional related information��� -=   $ stepsTaken_Grid��$��6��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� ->   $
 stepsTaken_Label��$��Steps taken:��$
��Left=�$��10,0,0,0q�$	��Top=�$^	 XList all steps taken during the call, including what effect each step had on the problem���� -?   $ stepsTaken_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $^	 XList all steps taken during the call, including what effect each step had on the problem��#��[  $( "Steps taken during troubleshooting��� -@   $ callResultStatus_Grid��$��7�� �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+�� -A   $ callResult_Grid��$��0��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -B   $
 callResult_Label��$��Result:��$��10,0,0,0q�$
��Left=�$	��Top=�$8	 2Choose a result to reflect the outcome of the call���� -C   $ callResult_ComboBox��$��	69,2,10,0q�$��Stretch=�$	��Top=�$8	 2Choose a result to reflect the outcome of the call��.@��� $��Resolved���� $��
Next steps���� $��	Escalated���� $��Transferred��� -D   $ resolution_Grid��$��1��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -E   $
 resolution_Label��$��Resolution:��$
��Left=�$��10,0,0,0q�$	��Top=�$Z	 TEnter what action resolved the issue or the answer that was provided to the customer���� -F   $ resolution_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$Z	 TEnter what action resolved the issue or the answer that was provided to the customer��#��[  $ What resolved the issue?��� -G   $ nextSteps_Grid��$��2��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -H   $
 nextSteps_Label��$��Next Steps:��$
��Left=�$��10,0,0,0q�$	��Top=�$7	 1Enter some recommended steps for the next advisor���� -I   $ nextSteps_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$7	 1Enter some recommended steps for the next advisor��#��[  $ Recommended next steps��� -J   $ escalation_Grid��$��3��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+�� -K   $ escalationType_Grid��$��0��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -L   $
 escalationType_Label��$��Escalated to:��$��10,0,0,0q�$
��Left=�$	��Top=�$<	 6Choose which department the case is being escalated to���� -M   $ escalationType_ComboBox��$��	99,2,10,0q�$��Stretch=�$	��Top=�$<	 6Choose which department the case is being escalated to��.@��� $#��Act! Premium Cloud Operations���� $��Emarketing - Technical���� $��Emarketing - Billing���� $��Emarketing - Cancellation���� $��Act! - Technical��� -N   $ cloudOpsEscalation_Grid��$��1��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+��� $��0��$��	20,0,10,0q�$��Stretch=�$	��Top=�$
��auto��$
��Wrap=��� -O   $D��>http://kb.internal.swiftpage.com/app/answers/detail/a_id/38247?�.�;9How to escalate a ticket to Act! Premium Cloud Operations� -P   $  cloudOpsEscalation_url_Row��$��1��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -Q   $"
 cloudOpsEscalation_url_Label��$
��URL:��$
��Left=�$��10,0,0,0q�$	��Top=�$5	 /Enter the affected Act! Premium Cloud Login URL���� -R   $$ cloudOpsEscalation_url_TextBox��$��Stretch=�$��
10,26,10,5q�$	��Top=�$
��auto�� . $5	 /Enter the affected Act! Premium Cloud Login URL��#��[  $ APC Login URL��� -S   $) #cloudOpsEscalation_databaseName_Row��$��2��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -T   $+
 %cloudOpsEscalation_databaseName_Label��$��	Database:��$
��Left=�$��10,0,0,0q�$	��Top=�$2	 ,Enter the database(s) affected by this issue���� -U   $- 'cloudOpsEscalation_databaseName_TextBox��$��Stretch=�$��
10,26,10,5q�$	��Top=�$
��auto�� . $2	 ,Enter the database(s) affected by this issue��#��[  $ Affected databases��� -V   $/ )cloudOpsEscalation_remoteDatabaseName_Row��$��3��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -W   $1
 +cloudOpsEscalation_remoteDatabaseName_Label��$
��RDB:��$
��Left=�$��10,0,0,0q�$	��Top=�$6	 0Enter the names of any affected remote databases���� -X   $3 -cloudOpsEscalation_remoteDatabaseName_TextBox��$��Stretch=�$��
10,26,10,5q�$	��Top=�$
��auto�� . $6	 0Enter the names of any affected remote databases��#��[  $ Affected remote databases��� -Y   $) #cloudOpsEscalation_PartnerName_Grid��$��4��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -Z   $*
 $cloudOpsEscalation_PartnerName_Label��$��Partner name:��$
��Left=�$��10,0,0,0q�$	��Top=�$a	 [If the APC subscription is paid for through a business partner, enter the partner name here���� -[   $, &cloudOpsEscalation_PartnerName_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$a	 [If the APC subscription is paid for through a business partner, enter the partner name here��#��[  $+ %Business Partner name (if applicable)��� -\   $, &cloudOpsEscalation_PartnerAccount_Grid��$��5��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -]   $-
 'cloudOpsEscalation_PartnerAccount_Label��$��Partner account:��$
��Left=�$��10,0,0,0q�$	��Top=�$m	 gIf the APC subscription is paid for through a business partner, enter the partner's account number here���� -^   $/ )cloudOpsEscalation_PartnerAccount_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$m	 gIf the APC subscription is paid for through a business partner, enter the partner's account number here��#��[  $5 /Business Partner account number (if applicable)��� -_   $+ %cloudOpsEscalation_AffectedUsers_Grid��$��6��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -`   $,
 &cloudOpsEscalation_AffectedUsers_Label��$��Affected users:��$
��Left=�$��10,0,0,0q�$	��Top=�$@	 :Enter all database users that are affected by this problem���� -a   $. (cloudOpsEscalation_AffectedUsers_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$@	 :Enter all database users that are affected by this problem��#��[  $ Affected database users��� -b   $1 +cloudOpsEscalation_PrimaryAccountEmail_Grid��$��7��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -c   $2
 ,cloudOpsEscalation_PrimaryAccountEmail_Label��$��Primary account email:��$
��Left=�$��10,0,0,0q�$	��Top=�$5	 /Email address that the account was created with���� -d   $4 .cloudOpsEscalation_PrimaryAccountEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$5	 /Email address that the account was created with��#��[  $# Primary account email address��� -e   $* $cloudOpsEscalation_ContactEmail_Grid��$��8��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -f   $+
 %cloudOpsEscalation_ContactEmail_Label��$��Contact email:��$
��Left=�$��10,0,0,0q�$	��Top=�$[	 UEmail address CloudOps should use to contact the customer relating to this escalation���� -g   $- 'cloudOpsEscalation_ContactEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$[	 UEmail address CloudOps should use to contact the customer relating to this escalation��#��[  $&  Customer's contact email address��� -h   $&  cloudOpsEscalation_Symptoms_Grid��$��9��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -i   $'
 !cloudOpsEscalation_Symptoms_Label��$��	Symptoms:��$
��Left=�$��10,0,0,0q�$	��Top=�$S	 MDetail the symptoms that the customer is seeing, including any error messages���� -j   $) #cloudOpsEscalation_Symptoms_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$S	 MDetail the symptoms that the customer is seeing, including any error messages��#��[  $) #What are the symptoms of the issue?��� -k   $, &cloudOpsEscalation_HowToReplicate_Grid��$��10��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -l   $-
 'cloudOpsEscalation_HowToReplicate_Label��$��How to replicate:��$
��Left=�$��10,0,0,0q�$	��Top=�$l	 fIf this is a recurring problem, enter some steps CloudOps can follow to reliably replicate the problem���� -m   $/ )cloudOpsEscalation_HowToReplicate_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$l	 fIf this is a recurring problem, enter some steps Cloudops can follow to reliably replicate the problem��#��[  $, &How do CloudOps replicate the problem?��� -n   $- 'cloudOpsEscalation_RequestedAction_Grid��$��11��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -o   $.
 (cloudOpsEscalation_RequestedAction_Label��$��Requested action:��$
��Left=�$��10,0,0,0q�$	��Top=�$�	 �Specify what action you need CloudOps to take, for example 'Investigate universal search issue' or 'Restore database with name: ...'���� -p   $0 *cloudOpsEscalation_RequestedAction_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$�	 �Specify what action you need CloudOps to take, for example 'Investigate universal search issue' or 'Restore database with name: ...'��#��[  $, &What actions do CloudOps need to take?��� -q   $( "emarketingTechnicalEscalation_Grid��$��2��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+��� $��0��$��	20,0,10,0q�$��Stretch=�$	��Top=�$
��auto��$
��Wrap=��� -r   $D��>http://kb.internal.swiftpage.com/app/answers/detail/a_id/36247?�.�31How to escalate a ticket to Emarketing T3 Support� -s   $2 ,emarketingTechnicalEscalation_AgentName_Grid��$��1��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -t   $3
 -emarketingTechnicalEscalation_AgentName_Label��$��Agent name:��$
��Left=�$��10,0,0,0q�$	��Top=�$;	 5Enter the name of the advisor creating the escalation���� -u   $5 /emarketingTechnicalEscalation_AgentName_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$;	 5Enter the name of the advisor creating the escalation��#��[  $ 	Your name��� -v   $4 .emarketingTechnicalEscalation_AccountName_Grid��$��2��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -w   $5
 /emarketingTechnicalEscalation_AccountName_Label��$��Account name:��$
��Left=�$��10,0,0,0q�$	��Top=�$D	 >Enter the Act! Emarketing or Swiftpage Emarketing account name���� -x   $7 1emarketingTechnicalEscalation_AccountName_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$D	 >Enter the Act! Emarketing or Swiftpage Emarketing account name��#��[  $ AEM or SPEM account name��� -y   $< 6emarketingTechnicalEscalation_CustomerPhoneNumber_Grid��$��3��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -z   $=
 7emarketingTechnicalEscalation_CustomerPhoneNumber_Label��$��Customer phone number:��$
��Left=�$��10,0,0,0q�$	��Top=�$z	 tEnter a phone number that Emarketing T3 support can use to contact the customer, including the country and area code���� -{   $? 9emarketingTechnicalEscalation_CustomerPhoneNumber_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$z	 tEnter a phone number that Emarketing T3 support can use to contact the customer, including the country and area code��#��[  $ Contact telephone number��� -|   $= 7emarketingTechnicalEscalation_CustomerEmailAddress_Grid��$��4��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -}   $>
 8emarketingTechnicalEscalation_CustomerEmailAddress_Label��$��Customer email address:��$
��Left=�$��10,0,0,0q�$	��Top=�$W	 QEnter an email address that Emarketing T3 support can use to contact the customer���� -~   $@ :emarketingTechnicalEscalation_CustomerEmailAddress_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$W	 QEnter an email address that Emarketing T3 support can use to contact the customer��#��[  $ Contact email address��� -   $/ )emarketingTechnicalEscalation_UserID_Grid��$��5��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $0
 *emarketingTechnicalEscalation_UserID_Label��$��User ID:��$
��Left=�$��10,0,0,0q�$	��Top=�$1	 +Enter the customer's User ID, if applicable���� -�   $2 ,emarketingTechnicalEscalation_UserID_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$1	 +Enter the customer's User ID, if applicable��#��[  $# Login User ID (if applicable)��� -�   $/ )emarketingTechnicalEscalation_Server_Grid��$��6��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $0
 *emarketingTechnicalEscalation_Server_Label��$��Server:��$
��Left=�$��10,0,0,0q�$	��Top=�$9	 3Enter the server that the customer is logging in to���� -�   $2 ,emarketingTechnicalEscalation_Server_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$9	 3Enter the server that the customer is logging in to��#��[  $ Server being logged in to��� -�   $4 .emarketingTechnicalEscalation_Integration_Grid��$��7��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $5
 /emarketingTechnicalEscalation_Integration_Label��$��Integration:��$
��Left=�$��10,0,0,0q�$	��Top=�$z	 tEnter the names and versions of any third party applications that the customer is using to integrate with Emarketing���� -�   $7 1emarketingTechnicalEscalation_Integration_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$z	 tEnter the names and versions of any third party applications that the customer is using to integrate with Emarketing��#��[  $* $Software integrating with Emarketing��� -�   $1 +emarketingTechnicalEscalation_Platform_Grid��$��8��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $2
 ,emarketingTechnicalEscalation_Platform_Label��$��	Platform:��$
��Left=�$��10,0,0,0q�$	��Top=�$4	 .Select which Emarketing platform is being used���� -�   $5 /emarketingTechnicalEscalation_Platform_ComboBox��$��
10,26,10,0q�$��Stretch=�$	��Top=�$4	 .Select which Emarketing platform is being used��.@��� $��Act! Emarketing���� $��Swiftpage Emarketing��� -�   $1 +emarketingTechnicalEscalation_Symptoms_Grid��$��9��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $2
 ,emarketingTechnicalEscalation_Symptoms_Label��$��	Symptoms:��$
��Left=�$��10,0,0,0q�$	��Top=�$S	 MDetail the symptoms that the customer is seeing, including any error messages���� -�   $4 .emarketingTechnicalEscalation_Symptoms_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$S	 MDetail the symptoms that the customer is seeing, including any error messages��#��[  $) #What are the symptoms of the issue?��� -�   $1 +emarketingTechnicalEscalation_DateTime_Grid��$��10��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $2
 ,emarketingTechnicalEscalation_DateTime_Label��$"��Date and time of occurrence:��$
��Left=�$��10,0,0,0q�$	��Top=�$�	 �Enter the date and time that the issue occurred, being as exact as possible. If the issue has occurred multiple times, enter each date and time.���� -�   $4 .emarketingTechnicalEscalation_DateTime_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$�	 �Enter the date and time that the issue occurred, being as exact as possible. If the issue has occurred multiple times, enter each date and time.��#��[  $ When did the issue occur?��� -�   $6 0emarketingTechnicalEscalation_ActionsAtTime_Grid��$��11��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $7
 1emarketingTechnicalEscalation_ActionsAtTime_Label��$(��"What happened when issue occurred:��$
��Left=�$��10,0,0,0q�$	��Top=�$O	 ISpecify what actions were being taken at the time that the issue occurred���� -�   $9 3emarketingTechnicalEscalation_ActionsAtTime_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$O	 ISpecify what actions were being taken at the time that the issue occurred��#��[  $2 ,What was being done when the issue occurred?��� -�   $5 /emarketingTechnicalEscalation_ReplicatedOn_Grid��$��12��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $6
 0emarketingTechnicalEscalation_ReplicatedOn_Label��$.��(Replicated with (account, user, server):��$
��Left=�$��10,0,0,0q�$	��Top=�$y	 sIf you were able to replicate the issue, specify the Account, Server, and User that you replicated the problem with���� -�   $8 2emarketingTechnicalEscalation_ReplicatedOn_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$y	 sIf you were able to replicate the issue, specify the Account, Server, and User that you replicated the problem with��#��[  $2 ,What account details did you replicate with?��� -�   $7 1emarketingTechnicalEscalation_HowToReplicate_Grid��$��13��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $8
 2emarketingTechnicalEscalation_HowToReplicate_Label��$��How to reproduce:��$
��Left=�$��10,0,0,0q�$	��Top=�$y	 sIf this is a recurring problem, enter some steps that Emarketing T3 support can use to reliably replicate the issue���� -�   $: 4emarketingTechnicalEscalation_HowToReplicate_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$y	 sIf this is a recurring problem, enter some steps that Emarketing T3 support can use to reliably replicate the issue��#��[  $' !How do you replicate the problem?��� -�   $8 2emarketingTechnicalEscalation_AccountOrGlobal_Grid��$��14��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $9
 3emarketingTechnicalEscalation_AccountOrGlobal_Label��$!��Global or account specific:��$
��Left=�$��10,0,0,0q�$	��Top=�$u	 oEnter the area that this issue has been isolated to - for example 'This account only', or 'All tested accounts'���� -�   $; 5emarketingTechnicalEscalation_AccountOrGlobal_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$u	 oEnter the area that this issue has been isolated to - for example 'This account only', or 'All tested accounts'��#��[  $$ Specific to this account only?��� -�   $&  emarketingBillingEscalation_Grid��$��3��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+��� $��0��$��	20,0,10,0q�$��Stretch=�$	��Top=�$
��auto��$
��Wrap=��� -�   $D��>http://kb.internal.swiftpage.com/app/answers/detail/a_id/38176?�.�/-How to submit a request to Emarketing Billing� -�   $2 ,emarketingBillingEscalation_AccountName_Grid��$��1��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $3
 -emarketingBillingEscalation_AccountName_Label��$��Account name:��$
��Left=�$��10,0,0,0q�$	��Top=�$D	 >Enter the Act! Emarketing or Swiftpage Emarketing account name���� -�   $5 /emarketingBillingEscalation_AccountName_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$D	 >Enter the Act! Emarketing or Swiftpage Emarketing account name��#��[  $ AEM or SPEM account name��� -�   $3 -emarketingBillingEscalation_AccountEmail_Grid��$��2��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $4
 .emarketingBillingEscalation_AccountEmail_Label��$��Account Email:��$
��Left=�$��10,0,0,0q�$	��Top=�$9	 3Enter the email address associated with the account���� -�   $6 0emarketingBillingEscalation_AccountEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$9	 3Enter the email address associated with the account��#��[  $ Account email address��� -�   $3 -emarketingBillingEscalation_ContactEmail_Grid��$��3��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $4
 .emarketingBillingEscalation_ContactEmail_Label��$��Contact Email:��$
��Left=�$��10,0,0,0q�$	��Top=�$E	 ?Enter an email address that can be used to contact the customer���� -�   $6 0emarketingBillingEscalation_ContactEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$E	 ?Enter an email address that can be used to contact the customer��#��[  $$ Customer contact email address��� -�   $3 -emarketingBillingEscalation_ContactPhone_Grid��$��4��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $4
 .emarketingBillingEscalation_ContactPhone_Label��$��Contact Phone:��$
��Left=�$��10,0,0,0q�$	��Top=�$h	 bEnter a phone number that can be used to contact the customer, including the country and area code���� -�   $6 0emarketingBillingEscalation_ContactPhone_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$h	 bEnter a phone number that can be used to contact the customer, including the country and area code��#��[  $# Customer contact phone number��� -�   $/ )emarketingBillingEscalation_Platform_Grid��$��5��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $0
 *emarketingBillingEscalation_Platform_Label��$��	Platform:��$
��Left=�$��10,0,0,0q�$	��Top=�$4	 .Select which Emarketing platform is being used���� -�   $3 -emarketingBillingEscalation_Platform_ComboBox��$��
10,26,10,0q�$��Stretch=�$	��Top=�$4	 .Select which Emarketing platform is being used��.@��� $��Act! Emarketing���� $��Swiftpage Emarketing��� -�   $: 4emarketingBillingEscalation_ReasonForEscalation_Grid��$��6��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $;
 5emarketingBillingEscalation_ReasonForEscalation_Label��$��Reason for escalation:��$
��Left=�$��10,0,0,0q�$	��Top=�$e	 _Enter some information on why this case is being escalated and what actions are being requested���� -�   $= 7emarketingBillingEscalation_ReasonForEscalation_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$e	 _Enter some information on why this case is being escalated and what actions are being requested��#��[  $) #Reason that case is being escalated��� -�   $+ %emarketingCancellationEscalation_Grid��$��4��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+��� $��0��$��	20,0,10,0q�$��Stretch=�$	��Top=�$
��auto��$
��Wrap=��� -�   $D��>http://kb.internal.swiftpage.com/app/answers/detail/a_id/36038?�.�*(How to cancel an Emarketing subscription� -�   $7 1emarketingCancellationEscalation_AccountName_Grid��$��1��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $8
 2emarketingCancellationEscalation_AccountName_Label��$��Account name:��$
��Left=�$��10,0,0,0q�$	��Top=�$D	 >Enter the Act! Emarketing or Swiftpage Emarketing account name���� -�   $: 4emarketingCancellationEscalation_AccountName_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$D	 >Enter the Act! Emarketing or Swiftpage Emarketing account name��#��[  $ AEM or SPEM account name��� -�   $8 2emarketingCancellationEscalation_AccountEmail_Grid��$��2��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $9
 3emarketingCancellationEscalation_AccountEmail_Label��$��Account email:��$
��Left=�$��10,0,0,0q�$	��Top=�$9	 3Enter the email address associated with the account���� -�   $; 5emarketingCancellationEscalation_AccountEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$9	 3Enter the email address associated with the account��#��[  $ Account email address��� -�   $D >emarketingCancellationEscalation_AccountOwnerContactEmail_Grid��$��3��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $E
 ?emarketingCancellationEscalation_AccountOwnerContactEmail_Label��$"��Account owner contact email:��$
��Left=�$��10,0,0,0q�$	��Top=�$I	 CEnter an email address that the account owner can be contacted with���� -�   $G AemarketingCancellationEscalation_AccountOwnerContactEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$I	 CEnter an email address that the account owner can be contacted with��#��[  $# Account owner's email address��� -�   $D >emarketingCancellationEscalation_AccountOwnerContactPhone_Grid��$��4��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $E
 ?emarketingCancellationEscalation_AccountOwnerContactPhone_Label��$"��Account owner contact phone:��$
��Left=�$��10,0,0,0q�$	��Top=�$m	 gEnter a phone number that the account owner can be contacted with, including the country and area codes���� -�   $G AemarketingCancellationEscalation_AccountOwnerContactPhone_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$m	 gEnter a phone number that the account owner can be contacted with, including the country and area codes��#��[  $" Account owner's phone number��� -�   $8 2emarketingCancellationEscalation_ContactEmail_Grid��$��5��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $9
 3emarketingCancellationEscalation_ContactEmail_Label��$��Contact email:��$
��Left=�$��10,0,0,0q�$	��Top=�$E	 ?Enter an email address that can be used to contact the customer���� -�   $; 5emarketingCancellationEscalation_ContactEmail_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$E	 ?Enter an email address that can be used to contact the customer��#��[  $$ Customer contact email address��� -�   $8 2emarketingCancellationEscalation_ContactPhone_Grid��$��6��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $9
 3emarketingCancellationEscalation_ContactPhone_Label��$��Contact phone:��$
��Left=�$��10,0,0,0q�$	��Top=�$i	 cEnter aa phone number that can be used to contact the customer, including the country and area code���� -�   $; 5emarketingCancellationEscalation_ContactPhone_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$h	 bEnter a phone number that can be used to contact the customer, including the country and area code��#��[  $# Customer contact phone number��� -�   $4 .emarketingCancellationEscalation_Platform_Grid��$��7��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $5
 /emarketingCancellationEscalation_Platform_Label��$��	Platform:��$
��Left=�$��10,0,0,0q�$	��Top=�$4	 .Select which Emarketing platform is being used���� -�   $8 2emarketingCancellationEscalation_Platform_ComboBox��$��
10,26,10,0q�$��Stretch=�$	��Top=�$4	 .Select which Emarketing platform is being used��.@��� $��Act! Emarketing���� $��Swiftpage Emarketing��� -�   $1 +emarketingCancellationEscalation_Issue_Grid��$��8��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $2
 ,emarketingCancellationEscalation_Issue_Label��$��Description of issue:��$
��Left=�$��10,0,0,0q�$	��Top=�$:	 4Enter a detailed description of the customer's issue���� -�   $4 .emarketingCancellationEscalation_Issue_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$:	 4Enter a detailed description of the customer's issue��#��[  $% Description of customer's issue��� -�   $? 9emarketingCancellationEscalation_ReasonForEscalation_Grid��$��9��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $@
 :emarketingCancellationEscalation_ReasonForEscalation_Label��$&�� Reason for cancellation request:��$
��Left=�$��10,0,0,0q�$	��Top=�$e	 _Enter some information on why this case is being escalated and what actions are being requested���� -�   $B <emarketingCancellationEscalation_ReasonForEscalation_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$e	 _Enter some information on why this case is being escalated and what actions are being requested��#��[  $) #Reason that case is being escalated��� -�   $ transfer_Grid��$��4��$y�	Collapsed=�$��0,0,0,0q�$��Stretch=�$
��auto�� �� $
��auto ��� $
��auto �.+�� -�   $ transferDepartment_Grid��$��0��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $
 transferDepartment_Label��$��Department:��$
��Left=�$��10,0,0,0q�$	��Top=�$A	 ;Enter the department that you are transferring this call to���� -�   $  transferDepartment_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$A	 ;Enter the department that you are transferring this call to��#��[  $# Department receiving transfer��� -�   $ transferPerson_Grid��$��1��$��0,0,0,0q�$��Stretch=�$
��auto��.+��� -�   $
 transferPerson_Label��$��Person:��$
��Left=�$��10,0,0,0q�$	��Top=�$I	 CEnter the name of the person that you are transferring this call to���� -�   $ transferPerson_TextBox��$��stretch=�$
��auto��$��
10,26,10,5q�$
 Wrap=�$	��Top=� . $
��Auto=�$I	 CEnter the name of the person that you are transferring this call to��#��[  $ Person receiving transfer��� -�   $  ticketLoggingButtonsBottom��$��4��$��Bottom=�$��43��.+��� $	��Top=�$��Stretch=�$��1�����	#FFFE5000�� -�   $ endTimer_Button��$��	End Timer��{�.  $��	10,0,0,10q�$
��Left=�$��Bottom=�$��70��$"	 Ends the call duration timer���� -�   $
 callDuration_Label��$��00:00:00��$��85,0,0,8q�$
��Left=�$��Bottom=��� -�   $ copyToClipboard_Button��$��Copy to Clipboard��$��Right=�$��	0,0,10,10q�$	��115��$��Bottom=�$��22��$R	 LCopies your notes to the clipboard ready to be pasted into your infor ticket��