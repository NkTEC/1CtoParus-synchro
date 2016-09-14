
sub xml_import
	Dim contragents '��������� XMLDOMNodeList ���� ��������� ��������� ����
	Dim cAgent '������� ��������� �����������
	Dim agncounter

	' ��������� XML-��������
	Set xmlParser = CreateObject("Msxml2.DOMDocument")
	xmlParser.async = False
	xmlParser.load "\\10.130.32.52\Tatneft\Mess_UH_20100.xml"

	' ��������� �� ������ ��������
	If xmlParser.parseError.errorCode Then
			MsgBox xmlParser.parseError.Reason
	End If

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.OpenTextFile("\\10.130.32.52\Tatneft\Exch_logs\1C-Parus_exchange.log", 8, True)
	MyFile.Write(vbNewLine&"**********************************************************************************************************************************"&vbNewLine&vbTab&now()&vbTab&" ������ ������� ������ �� 1�:�� � ��� �����"&vbNewLine)

	' ������� ���� �����������
	Set contragents = xmlParser.selectNodes("//�����������/������")
	If contragents.length > 0 then
		' ���������� ������ ������������ � XML-���������
		object_counter=0
		T_Agn_error = False
		For Each nodeNode In contragents
			SAPcode				= NULL
			INN					= NULL
			KPP					= NULL
			agntype				= NULL
			agncounter			= NULL
			DoubledAgentsString	= NULL
			SAPcode				= NULL
			newRN				= NULL
			
			object_counter=object_counter+1
			
			SAPcode = RTrim(nodeNode.selectSingleNode("���").text)
			INN = nodeNode.selectSingleNode("���").text
			KPP = nodeNode.selectSingleNode("���").text
			
			if INN="1651057954" then
				divider="/"
			else
				divider="^"
			end if

			If nodeNode.selectSingleNode("�������������������������").text = "���������������" Then
					agntype=0
			Else
					agntype=1
			end If

			' ������ RN ����������� �� ���� SAP � ������� ������� ��������
			Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='"&SAPcode&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - ��� �������� "��� SAP" (1�������� ����� �������� � ���������)
			Query.Open
			If Query.IsEmpty then

				' ����� � ��� ��� ���� ���������� � ����� �� ���/��� �� ��� ���� SAP
				If KPP="" then
					CaQuery=Query
					CaQuery.Sql.Text = "select AGNNAME, AGNIDNUMB, REASON_CODE from AGNLIST where AGNIDNUMB='"&INN&"'"
					CaQuery.Open
				else
					CaQuery=Query
					CaQuery.Sql.Text = "select AGNNAME, AGNIDNUMB, REASON_CODE from AGNLIST where AGNIDNUMB='"&INN&"' and REASON_CODE='"&KPP&"'"
					CaQuery.Open
				end if

				If CaQuery.IsEmpty then
					MyFile.Write("INFO "&now()&vbTab&" � ����� �� ������ ���������� � ����� SAP "&SAPcode&" ("&nodeNode.selectSingleNode("������������").text&") - ������ ������ �����������."&vbNewLine)
					
					'������� ������ � ����� �����������
					StoredProc.StoredProcName="P_AGNLIST_INSERT"
					StoredProc.ParamByName("nCOMPANY").value=42903                                        '��� �������������
					StoredProc.ParamByName("CRN").value=155332                                                '��� ��������
					StoredProc.ParamByName("AGNTYPE").value=agntype
					StoredProc.ParamByName("AGNNAME").value=nodeNode.selectSingleNode("������������").text
					StoredProc.ParamByName("sFULLNAME").value=nodeNode.selectSingleNode("������������������").text
					StoredProc.ParamByName("AGNIDNUMB").value=INN
					StoredProc.ParamByName("sREASON_CODE").value=nodeNode.selectSingleNode("���").text
					StoredProc.ParamByName("sOGRN").value=nodeNode.selectSingleNode("����").text
					StoredProc.ParamByName("ORGCODE").value=nodeNode.selectSingleNode("����").text
					StoredProc.ParamByName("AGNABBR").value=INN&divider&KPP
					StoredProc.ParamByName("PHONE").value=nodeNode.selectSingleNode("�������").text
					StoredProc.ParamByName("EMP").value=0
					StoredProc.ParamByName("nSEX").value=0
					StoredProc.ParamByName("nRESIDENT_SIGN").value=0
					StoredProc.ParamByName("nCOEFFIC").value=0
					StoredProc.ExecProc
					newRN = StoredProc.ParamByName("nRN").value

					' ������� ��� SAP � �������� �������
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" '���������� ��� ������ ��� ��������
					StoredProc.ParamByName("PROPERTY").value="��������1�"
					StoredProc.ParamByName("UNITCODE").value="AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value=newRN
					StoredProc.ParamByName("ST_VAL").value=SAPcode
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc

					'������� ������ �� ������ ������ �����������
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" '���������� ��� ������ ��� ��������
					StoredProc.ParamByName("PROPERTY").value="������������1�"
					StoredProc.ParamByName("UNITCODE").value="AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value=newRN
					StoredProc.ParamByName("ST_VAL").value=nodeNode.selectSingleNode("����������������").text
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" '���������� ��� ������ ��� ��������
					StoredProc.ParamByName("PROPERTY").value="������������1�"
					StoredProc.ParamByName("UNITCODE").value="AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value=newRN
					StoredProc.ParamByName("ST_VAL").value=nodeNode.selectSingleNode("�������������").text
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc
				else
					' ���/��� ���������, � ���� SAP ��� - �� �������
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" � ������� ������ ���������� � ������ �� ���/���, �� ��� ���� SAP "&SAPcode&" - ���������� �� ������, ���������� ��������� �������: "&INN&"/"&KPP&" ("&CaQuery.FieldByname("AGNNAME").value&")"&vbNewLine)
					T_Agn_error = True
				end If
				CaQuery.Close
			else        '������� ������ � ����� ����� SAP

				' ������ ����������� � ������� �� RN
				CaQuery=Query
				CaQuery.Sql.Text = "select * from AGNLIST where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				CaQuery.Open

				' ���������� ���������� ��������� ������������ � �������� �� ������ �� ������ ������
				agncounter = 0
				DoubledAgentsString=""
				do while not CaQuery.EOF
					agncounter=agncounter+1
					DoubledAgentsString=DoubledAgentsString & vbTab&agncounter & ") " & CaQuery.FieldByname("AGNIDNUMB").value & "/" & CaQuery.FieldByname("REASON_CODE").value & " " & CaQuery.FieldByname("AGNNAME").value & vbNewLine
					CaQuery.next
				loop
				If agncounter>1 then
					'�������� �� ������������ ����� SAP � ������� ������� � ����
					MyFile.Write("ERROR "&now()&vbTab&" � ����� ������� ��������� ������������ � ����� SAP "&SAPcode&" ("&nodeNode.selectSingleNode("������������").text&"),  - ���������� ��������� �������, ������ �� ��������:"&vbNewLine)
					T_Agn_error = True
					MyFile.Write(DoubledAgentsString)
				elseIf agncounter=1 Then
					MyFile.Write("INFO "&now()&vbTab&" � ����� ������ ���������� � ����� SAP "&SAPcode&" ("&nodeNode.selectSingleNode("������������").text&") - �������� ������ �����������."&vbNewLine)

					
					RN					= CaQuery.FieldByname("RN").value
					ECONCODE			= CaQuery.FieldByname("ECONCODE").value
					If nodeNode.selectSingleNode("����").text = "" then
						ORGCODE			= CaQuery.FieldByname("ORGCODE").value
					else
						ORGCODE			= nodeNode.selectSingleNode("����").text
					end if
					AGNFAMILYNAME		= CaQuery.FieldByname("AGNFAMILYNAME").value
					AGNFIRSTNAME		= CaQuery.FieldByname("AGNFIRSTNAME").value
					AGNLASTNAME			= CaQuery.FieldByname("AGNLASTNAME").value
					AGNFAMILYNAME_TO	= CaQuery.FieldByname("AGNFAMILYNAME_TO").value
					AGNFIRSTNAME_TO		= CaQuery.FieldByname("AGNFIRSTNAME_TO").value
					AGNLASTNAME_TO		= CaQuery.FieldByname("AGNLASTNAME_TO").value
					AGNFAMILYNAME_FR	= CaQuery.FieldByname("AGNFAMILYNAME_FR").value
					AGNFIRSTNAME_FR		= CaQuery.FieldByname("AGNFIRSTNAME_FR").value
					AGNLASTNAME_FR		= CaQuery.FieldByname("AGNLASTNAME_FR").value
					AGNFAMILYNAME_AC	= CaQuery.FieldByname("AGNFAMILYNAME_AC").value
					AGNFIRSTNAME_AC		= CaQuery.FieldByname("AGNFIRSTNAME_AC").value
					AGNLASTNAME_AC		= CaQuery.FieldByname("AGNLASTNAME_AC").value
					AGNFAMILYNAME_ABL	= CaQuery.FieldByname("AGNFAMILYNAME_ABL").value
					AGNFIRSTNAME_ABL	= CaQuery.FieldByname("AGNFIRSTNAME_ABL").value
					AGNLASTNAME_ABL		= CaQuery.FieldByname("AGNLASTNAME_ABL").value
					EMPPOST				= CaQuery.FieldByname("EMPPOST").value
					EMPPOST_FROM		= CaQuery.FieldByname("EMPPOST_FROM").value
					EMPPOST_TO			= CaQuery.FieldByname("EMPPOST_TO").value
					EMPPOST_AC			= CaQuery.FieldByname("EMPPOST_AC").value
					EMPPOST_ABL			= CaQuery.FieldByname("EMPPOST_ABL").value
					AGNBURN				= CaQuery.FieldByname("AGNBURN").value
					If nodeNode.selectSingleNode("�������").text = "" then
						PHONE			= CaQuery.FieldByname("PHONE").value
					else
						PHONE			= nodeNode.selectSingleNode("�������").text
					end if
					PHONE2				= CaQuery.FieldByname("PHONE2").value
					FAX					= CaQuery.FieldByname("FAX").value
					TELEX				= CaQuery.FieldByname("TELEX").value
					If nodeNode.selectSingleNode("����������������").text = "" then
						MAIL			= CaQuery.FieldByname("MAIL").value
					else
						MAIL			= nodeNode.selectSingleNode("����������������").text
					end if
					IMAGE				= CaQuery.FieldByname("IMAGE").value
					DISCDATE			= CaQuery.FieldByname("DISCDATE").value
					AGN_COMMENT			= CaQuery.FieldByname("AGN_COMMENT").value
					PENSION_NBR			= CaQuery.FieldByname("PENSION_NBR").value
					MEDPOLICY_SER		= CaQuery.FieldByname("MEDPOLICY_SER").value
					MEDPOLICY_NUMB		= CaQuery.FieldByname("MEDPOLICY_NUMB").value
					PROPFORM			= CaQuery.FieldByname("PROPFORM").value
					TAXPSTATUS			= CaQuery.FieldByname("TAXPSTATUS").value
					PRFMLSTS			= CaQuery.FieldByname("PRFMLSTS").value
					PRNATION			= CaQuery.FieldByname("PRNATION").value
					CITIZENSHIP			= CaQuery.FieldByname("CITIZENSHIP").value
					CITIZENOKIN			= CaQuery.FieldByname("CITIZENOKIN").value
					ADDR_BURN			= CaQuery.FieldByname("ADDR_BURN").value
					PRMLREL				= CaQuery.FieldByname("PRMLREL").value
					OKATO				= CaQuery.FieldByname("OKATO").value
					PFR_NAME			= CaQuery.FieldByname("PFR_NAME").value
					PFR_FILL_DATE		= CaQuery.FieldByname("PFR_FILL_DATE").value
					PFR_REG_DATE		= CaQuery.FieldByname("PFR_REG_DATE").value
					PFR_REG_NUMB		= CaQuery.FieldByname("PFR_REG_NUMB").value
					If nodeNode.selectSingleNode("����").text = "" then
						OGRN			= CaQuery.FieldByname("OGRN").value
					else
						OGRN			= nodeNode.selectSingleNode("����").text
					end if
					OKFS				= CaQuery.FieldByname("OKFS").value
					If nodeNode.selectSingleNode("�����").text = "" then
						OKOPF			= CaQuery.FieldByname("OKOPF").value
					else
						OKOPF			= nodeNode.selectSingleNode("�����").text
					end if
					TFOMS				= CaQuery.FieldByname("TFOMS").value
					FSS_REG_NUMB		= CaQuery.FieldByname("FSS_REG_NUMB").value
					FSS_SUBCODE			= CaQuery.FieldByname("FSS_SUBCODE").value
					AGNDEATH			= CaQuery.FieldByname("AGNDEATH").value
					OKTMO				= CaQuery.FieldByname("OKTMO").value
					INN_CITIZENSHIP		= CaQuery.FieldByname("INN_CITIZENSHIP").value
					
					'��������� � NULL ��� �������� 0
					If PROPFORM = 0 then
						propform = NULL
					else
						Query.SQL.Text = "select CODE from PROPFORMS where RN='"&PROPFORM&"'"
						Query.Open
						propform = Query.FieldByname("CODE").value
					end if			
					
					If TAXPSTATUS = 0 then
						taxpstatus = NULL
					else
						Query.SQL.Text = "select CODE from TAXPAYERSTATUS where RN='"&TAXPSTATUS&"'"
						Query.Open
						TAXPSTATUS = Query.FieldByname("CODE").value
					end if				
					
					If PRFMLSTS = 0 then
						prfmlsts = NULL
					end if	
					
					If PRNATION = 0 then
						prnation = NULL
					end if				
					
					If CITIZENSHIP = 0 then
						citizenship = NULL
					end if	
					
					If PRMLREL = 0 then
						prmlrel = NULL
					end if					
					
					If OKATO = 0 then
						OKATO = NULL
					end if					
					If OKTMO = 0 then
						OKTMO = NULL
					end if
										
					'������� ������ � �����������
					StoredProc.StoredProcName="P_AGNLIST_UPDATE"
					StoredProc.ParamByName("nCOMPANY").value		= 42903	'��� �������������
					StoredProc.ParamByName("RN").value				= RN
					StoredProc.ParamByName("AGNABBR").value			= nodeNode.selectSingleNode("���").text&divider&nodeNode.selectSingleNode("���").text
					StoredProc.ParamByName("AGNTYPE").value			= agntype
					StoredProc.ParamByName("AGNNAME").value			= nodeNode.selectSingleNode("������������").text
					StoredProc.ParamByName("AGNIDNUMB").value		= nodeNode.selectSingleNode("���").text
					StoredProc.ParamByName("ECONCODE").value		= ECONCODE
					StoredProc.ParamByName("ORGCODE").value			= nodeNode.selectSingleNode("����").text
					StoredProc.ParamByName("AGNFAMILYNAME").value	= AGNFAMILYNAME
					StoredProc.ParamByName("AGNFIRSTNAME").value	= AGNFIRSTNAME
					StoredProc.ParamByName("AGNLASTNAME").value		= AGNLASTNAME
					StoredProc.ParamByName("AGNFAMILYNAME_TO").value= AGNFAMILYNAME_TO
					StoredProc.ParamByName("AGNFIRSTNAME_TO").value	= AGNFIRSTNAME_TO
					StoredProc.ParamByName("AGNLASTNAME_TO").value	= AGNLASTNAME_TO
					StoredProc.ParamByName("AGNFAMILYNAME_FR").value= AGNFAMILYNAME_FR
					StoredProc.ParamByName("AGNFIRSTNAME_FR").value	= AGNFIRSTNAME_FR
					StoredProc.ParamByName("AGNLASTNAME_FR").value	= AGNLASTNAME_FR
					StoredProc.ParamByName("AGNFAMILYNAME_AC").value= AGNFAMILYNAME_AC
					StoredProc.ParamByName("AGNFIRSTNAME_AC").value	= AGNFIRSTNAME_AC
					StoredProc.ParamByName("AGNLASTNAME_AC").value	= AGNLASTNAME_AC
					StoredProc.ParamByName("AGNFAMILYNAME_ABL").value= AGNFAMILYNAME_ABL
					StoredProc.ParamByName("AGNFIRSTNAME_ABL").value= AGNFIRSTNAME_ABL
					StoredProc.ParamByName("AGNLASTNAME_ABL").value	= AGNLASTNAME_ABL
					StoredProc.ParamByName("EMP").value				= 0
					StoredProc.ParamByName("EMPPOST").value			= EMPPOST
					StoredProc.ParamByName("EMPPOST_FROM").value	= EMPPOST_FROM
					StoredProc.ParamByName("EMPPOST_TO").value		= EMPPOST_TO
					StoredProc.ParamByName("EMPPOST_AC").value		= EMPPOST_AC
					StoredProc.ParamByName("EMPPOST_ABL").value		= EMPPOST_ABL
					StoredProc.ParamByName("AGNBURN").value			= AGNBURN
					StoredProc.ParamByName("PHONE").value			= PHONE
					StoredProc.ParamByName("PHONE2").value			= PHONE2
					StoredProc.ParamByName("FAX").value				= FAX
					StoredProc.ParamByName("TELEX").value			= TELEX
					StoredProc.ParamByName("MAIL").value			= MAIL
					StoredProc.ParamByName("IMAGE").value			= IMAGE
					StoredProc.ParamByName("dDISCDATE").value		= DISCDATE
					StoredProc.ParamByName("AGN_COMMENT").value		= AGN_COMMENT
					StoredProc.ParamByName("nSEX").value			= 0
					StoredProc.ParamByName("sPENSION_NBR").value	= PENSION_NBR
					StoredProc.ParamByName("sMEDPOLICY_SER").value	= MEDPOLICY_SER
					StoredProc.ParamByName("sMEDPOLICY_NUMB").value	= MEDPOLICY_NUMB
					StoredProc.ParamByName("sPROPFORM").value		= propform
					StoredProc.ParamByName("sREASON_CODE").value	= nodeNode.selectSingleNode("���").text
					StoredProc.ParamByName("nRESIDENT_SIGN").value	= 0
					StoredProc.ParamByName("sTAXPSTATUS").value		= taxpstatus
					StoredProc.ParamByName("sOGRN").value			= OGRN
					StoredProc.ParamByName("sPRFMLSTS").value		= prfmlsts
					StoredProc.ParamByName("sPRNATION").value		= prnation
					StoredProc.ParamByName("sCITIZENSHIP").value	= citizenship
					StoredProc.ParamByName("CITIZENOKIN").value		= CITIZENOKIN
					StoredProc.ParamByName("ADDR_BURN").value		= ADDR_BURN
					StoredProc.ParamByName("sPRMLREL").value		= prmlrel
					StoredProc.ParamByName("sOKATO").value			= OKATO
					StoredProc.ParamByName("sPFR_NAME").value		= PFR_NAME
					StoredProc.ParamByName("dPFR_FILL_DATE").value	= PFR_FILL_DATE
					StoredProc.ParamByName("dPFR_REG_DATE").value	= PFR_REG_DATE
					StoredProc.ParamByName("sPFR_REG_NUMB").value	= PFR_REG_NUMB
					StoredProc.ParamByName("sFULLNAME").value		= nodeNode.selectSingleNode("������������������").text
					StoredProc.ParamByName("sOKFS").value			= OKFS
					StoredProc.ParamByName("sOKOPF").value			= OKOPF
					StoredProc.ParamByName("sTFOMS").value			= TFOMS
					StoredProc.ParamByName("sFSS_REG_NUMB").value	= FSS_REG_NUMB
					StoredProc.ParamByName("sFSS_SUBCODE").value	= FSS_SUBCODE
					StoredProc.ParamByName("nCOEFFIC").value		= 0
					StoredProc.ParamByName("dAGNDEATH").value		= AGNDEATH
					StoredProc.ParamByName("sOKTMO").value			= OKTMO
					StoredProc.ParamByName("sINN_CITIZENSHIP").value= INN_CITIZENSHIP
					StoredProc.ExecProc

					'������� ������ �� ������ �����������
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" '���������� ��� ������ ��� ��������
					StoredProc.ParamByName("PROPERTY").value	= "������������1�"
					StoredProc.ParamByName("UNITCODE").value	= "AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value		= RN
					StoredProc.ParamByName("ST_VAL").value		= nodeNode.selectSingleNode("����������������").text
					StoredProc.ParamByName("NUM_VAL").value		= NULL
					StoredProc.ExecProc
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" '���������� ��� ������ ��� ��������
					StoredProc.ParamByName("PROPERTY").value	= "������������1�"
					StoredProc.ParamByName("UNITCODE").value	= "AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value		= RN
					StoredProc.ParamByName("ST_VAL").value		= nodeNode.selectSingleNode("�������������").text
					StoredProc.ParamByName("NUM_VAL").value		= NULL
					StoredProc.ExecProc
				end If
				CaQuery.Close
			end if
			Query.Close
		Next
		MyFile.Write("INFO "&now()&vbTab&" ����������� ���������. ����� ����������: "& object_counter & vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" � ������� �������� �� ������� ���������� � ������������."&vbNewLine)
	End if
	
	If T_Agn_error = True then
		MsgBox "�������� ������������ ������ � �������. ��������� ������ 1C-Parus_exchange.log"
		Exit Sub
	end if

	'������� ���� "�����"
	Set Banks = xmlParser.selectNodes("//�����/������")
	If Banks.length > 0 then
		'���������� ������ ������
		object_counter=0
		For Each nodeNode In Banks
			object_counter=object_counter+1
			agnbank_rn	= NULL
			
			BankQuery=Query
			BankQuery.Sql.Text = "select RN, AGNRN from AGNBANKS where BANKFCODEACC='"&nodeNode.selectSingleNode("���").text&"' and CRN='104583471'" '���� ����� ������ � ��������, �������� ��������� ��� ������ �� 1�
			BankQuery.Open
			agnbank_rn=BankQuery.FieldByname("RN").value
			If BankQuery.IsEmpty Then
				MyFile.Write("INFO "&now()&vbTab&" � ����� �� ������ ���� � ����� ��� "&nodeNode.selectSingleNode("���").text&" ("&nodeNode.selectSingleNode("������������").text&") - ������ ����� ����."&vbNewLine)
				'������� ������ � ����� �����������-�����
				StoredProc.StoredProcName="P_AGNLIST_INSERT"
				StoredProc.ParamByName("nCOMPANY").value=42903                                        '��� �������������
				StoredProc.ParamByName("CRN").value=104582949                                        '��� �������� � �������
				StoredProc.ParamByName("AGNTYPE").value=0
				StoredProc.ParamByName("AGNNAME").value=nodeNode.selectSingleNode("������������").text&", "&nodeNode.selectSingleNode("�����").text
				StoredProc.ParamByName("AGNABBR").value="����_"&nodeNode.selectSingleNode("���").text
				StoredProc.ParamByName("EMP").value=0
				StoredProc.ParamByName("nSEX").value=0
				StoredProc.ParamByName("nRESIDENT_SIGN").value=0
				StoredProc.ParamByName("nCOEFFIC").value=0
				StoredProc.ExecProc

				'������� ������ � ����� �������� ������� "���������� ����������"
				StoredProc.StoredProcName="P_AGNBANKS_INSERT"
				StoredProc.ParamByName("nCOMPANY").value=42903                                        '��� �������������
				StoredProc.ParamByName("nCRN").value=104583471                                        '��� �������� � �������
				StoredProc.ParamByName("sBANKFCODEACC").value=nodeNode.selectSingleNode("���").text
				StoredProc.ParamByName("sBANKACC").value=nodeNode.selectSingleNode("��������").text
				StoredProc.ParamByName("sCODE").value="����_"&nodeNode.selectSingleNode("���").text
				StoredProc.ExecProc
			else
				MyFile.Write("INFO "&now()&vbTab&" � ����� ������ ���� � ����� ��� "&nodeNode.selectSingleNode("���").text&" ("&nodeNode.selectSingleNode("������������").text&") - �������� ������������ ����."&vbNewLine)
				bAgnQuery = Query
				bAgnQuery.SQL.Text="select * from agnlist where RN='"&BankQuery.FieldByname("AGNRN").value&"'"
				bAgnQuery.Open

				REM '������� ������ � �����������-�����
				REM StoredProc.StoredProcName="P_AGNLIST_UPDATE"
				REM StoredProc.ParamByName("RN").value=bAgnQuery.FieldByname("RN").value                                                '��� �����������
				REM StoredProc.ParamByName("AGNABBR").value="����_"&nodeNode.selectSingleNode("���").text
				REM StoredProc.ParamByName("AGNTYPE").value=0
				REM StoredProc.ParamByName("AGNNAME").value=nodeNode.selectSingleNode("������������").text&", "&nodeNode.selectSingleNode("�����").text
				REM StoredProc.ParamByName("EMP").value=0
				REM StoredProc.ParamByName("nSEX").value=0
				REM StoredProc.ParamByName("nRESIDENT_SIGN").value=0
				REM StoredProc.ParamByName("nCOEFFIC").value=0
				REM StoredProc.ExecProc
				REM bAgnQuery.Close
				
				'��������� � NULL ��� �������� 0
				If Query.FieldByname("PROPFORM").value = 0 then
					propform = NULL
				else
					propform = Query.FieldByname("PROPFORM").value
				end if				
				If Query.FieldByname("TAXPSTATUS").value = 0 then
					taxpstatus = NULL
				else
					taxpstatus = Query.FieldByname("TAXPSTATUS").value
				end if				
				If Query.FieldByname("PRFMLSTS").value = 0 then
					prfmlsts = NULL
				else
					prfmlsts = Query.FieldByname("PRFMLSTS").value
				end if				
				If Query.FieldByname("PRNATION").value = 0 then
					prnation = NULL
				else
					prnation = Query.FieldByname("PRNATION").value
				end if				
				If Query.FieldByname("CITIZENSHIP").value = 0 then
					citizenship = NULL
				else
					citizenship = Query.FieldByname("CITIZENSHIP").value
				end if
				If Query.FieldByname("PRMLREL").value = 0 then
					prmlrel = NULL
				else
					prmlrel = Query.FieldByname("PRMLREL").value
				end if				
				If Query.FieldByname("OKATO").value = 0 then
					OKATO = NULL
				else
					OKATO = Query.FieldByname("OKATO").value
				end if				
				If Query.FieldByname("OKTMO").value = 0 then
					OKTMO = NULL
				else
					OKTMO = Query.FieldByname("OKTMO").value
				end if
				
				StoredProc.StoredProcName="P_AGNLIST_UPDATE"
				StoredProc.ParamByName("nCOMPANY").value		= 42903	'��� �������������
				StoredProc.ParamByName("RN").value				= bAgnQuery.FieldByname("RN").value	'��� �����������
				StoredProc.ParamByName("AGNABBR").value			= "����_"&nodeNode.selectSingleNode("���").text
				StoredProc.ParamByName("AGNTYPE").value			= 0
				StoredProc.ParamByName("AGNNAME").value			= nodeNode.selectSingleNode("������������").text&", "&nodeNode.selectSingleNode("�����").text
				StoredProc.ParamByName("AGNIDNUMB").value		= bAgnQuery.FieldByname("AGNIDNUMB").value
				StoredProc.ParamByName("ECONCODE").value		= bAgnQuery.FieldByname("ECONCODE").value
				StoredProc.ParamByName("ORGCODE").value			= bAgnQuery.FieldByname("ORGCODE").value
				StoredProc.ParamByName("AGNFAMILYNAME").value	= bAgnQuery.FieldByname("AGNFAMILYNAME").value
				StoredProc.ParamByName("AGNFIRSTNAME").value	= bAgnQuery.FieldByname("AGNFIRSTNAME").value
				StoredProc.ParamByName("AGNLASTNAME").value		= bAgnQuery.FieldByname("AGNLASTNAME").value
				StoredProc.ParamByName("AGNFAMILYNAME_TO").value= bAgnQuery.FieldByname("AGNFAMILYNAME_TO").value
				StoredProc.ParamByName("AGNFIRSTNAME_TO").value	= bAgnQuery.FieldByname("AGNFIRSTNAME_TO").value
				StoredProc.ParamByName("AGNLASTNAME_TO").value	= bAgnQuery.FieldByname("AGNLASTNAME_TO").value
				StoredProc.ParamByName("AGNFAMILYNAME_FR").value= bAgnQuery.FieldByname("AGNFAMILYNAME_FR").value
				StoredProc.ParamByName("AGNFIRSTNAME_FR").value	= bAgnQuery.FieldByname("AGNFIRSTNAME_FR").value
				StoredProc.ParamByName("AGNLASTNAME_FR").value	= bAgnQuery.FieldByname("AGNLASTNAME_FR").value
				StoredProc.ParamByName("AGNFAMILYNAME_AC").value= bAgnQuery.FieldByname("AGNFAMILYNAME_AC").value
				StoredProc.ParamByName("AGNFIRSTNAME_AC").value	= bAgnQuery.FieldByname("AGNFIRSTNAME_AC").value
				StoredProc.ParamByName("AGNLASTNAME_AC").value	= bAgnQuery.FieldByname("AGNLASTNAME_AC").value
				StoredProc.ParamByName("AGNFAMILYNAME_ABL").value= bAgnQuery.FieldByname("AGNFAMILYNAME_ABL").value
				StoredProc.ParamByName("AGNFIRSTNAME_ABL").value= bAgnQuery.FieldByname("AGNFIRSTNAME_ABL").value
				StoredProc.ParamByName("AGNLASTNAME_ABL").value	= bAgnQuery.FieldByname("AGNLASTNAME_ABL").value
				StoredProc.ParamByName("EMP").value				= 0
				StoredProc.ParamByName("EMPPOST").value			= bAgnQuery.FieldByname("AGNLASTNAME_ABL").value
				StoredProc.ParamByName("EMPPOST_FROM").value	= bAgnQuery.FieldByname("EMPPOST_FROM").value
				StoredProc.ParamByName("EMPPOST_TO").value		= bAgnQuery.FieldByname("EMPPOST_TO").value
				StoredProc.ParamByName("EMPPOST_AC").value		= bAgnQuery.FieldByname("EMPPOST_AC").value
				StoredProc.ParamByName("EMPPOST_ABL").value		= bAgnQuery.FieldByname("EMPPOST_ABL").value
				StoredProc.ParamByName("AGNBURN").value			= bAgnQuery.FieldByname("AGNBURN").value
				StoredProc.ParamByName("PHONE").value			= bAgnQuery.FieldByname("PHONE").value
				StoredProc.ParamByName("PHONE2").value			= bAgnQuery.FieldByname("PHONE2").value
				StoredProc.ParamByName("FAX").value				= bAgnQuery.FieldByname("FAX").value
				StoredProc.ParamByName("TELEX").value			= bAgnQuery.FieldByname("TELEX").value
				StoredProc.ParamByName("MAIL").value			= bAgnQuery.FieldByname("MAIL").value
				StoredProc.ParamByName("IMAGE").value			= bAgnQuery.FieldByname("IMAGE").value
				StoredProc.ParamByName("dDISCDATE").value		= bAgnQuery.FieldByname("DISCDATE").value
				StoredProc.ParamByName("AGN_COMMENT").value		= bAgnQuery.FieldByname("AGN_COMMENT").value
				StoredProc.ParamByName("nSEX").value			= 0
				StoredProc.ParamByName("sPENSION_NBR").value	= bAgnQuery.FieldByname("PENSION_NBR").value
				StoredProc.ParamByName("sMEDPOLICY_SER").value	= bAgnQuery.FieldByname("MEDPOLICY_SER").value
				StoredProc.ParamByName("sMEDPOLICY_NUMB").value	= bAgnQuery.FieldByname("MEDPOLICY_NUMB").value
				StoredProc.ParamByName("sPROPFORM").value		= propform
				StoredProc.ParamByName("sREASON_CODE").value	= bAgnQuery.FieldByname("REASON_CODE").value
				StoredProc.ParamByName("nRESIDENT_SIGN").value	= 0
				StoredProc.ParamByName("sTAXPSTATUS").value		= taxpstatus
				StoredProc.ParamByName("sOGRN").value			= bAgnQuery.FieldByname("OGRN").value
				StoredProc.ParamByName("sPRFMLSTS").value		= prfmlsts
				StoredProc.ParamByName("sPRNATION").value		= prnation
				StoredProc.ParamByName("sCITIZENSHIP").value	= citizenship
				StoredProc.ParamByName("CITIZENOKIN").value		= bAgnQuery.FieldByname("CITIZENOKIN").value
				StoredProc.ParamByName("ADDR_BURN").value		= bAgnQuery.FieldByname("ADDR_BURN").value
				StoredProc.ParamByName("sPRMLREL").value		= prmlrel
				StoredProc.ParamByName("sOKATO").value			= OKATO
				StoredProc.ParamByName("sPFR_NAME").value		= bAgnQuery.FieldByname("PFR_NAME").value
				StoredProc.ParamByName("dPFR_FILL_DATE").value	= bAgnQuery.FieldByname("PFR_FILL_DATE").value
				StoredProc.ParamByName("dPFR_REG_DATE").value	= bAgnQuery.FieldByname("PFR_REG_DATE").value
				StoredProc.ParamByName("sPFR_REG_NUMB").value	= bAgnQuery.FieldByname("PFR_REG_NUMB").value
				StoredProc.ParamByName("sFULLNAME").value		= bAgnQuery.FieldByname("FULLNAME").value
				StoredProc.ParamByName("sOKFS").value			= bAgnQuery.FieldByname("OKFS").value
				StoredProc.ParamByName("sOKOPF").value			= bAgnQuery.FieldByname("OKOPF").value
				StoredProc.ParamByName("sTFOMS").value			= bAgnQuery.FieldByname("TFOMS").value
				StoredProc.ParamByName("sFSS_REG_NUMB").value	= bAgnQuery.FieldByname("FSS_REG_NUMB").value
				StoredProc.ParamByName("sFSS_SUBCODE").value	= bAgnQuery.FieldByname("FSS_SUBCODE").value
				StoredProc.ParamByName("nCOEFFIC").value		= 0
				StoredProc.ParamByName("dAGNDEATH").value		= bAgnQuery.FieldByname("AGNDEATH").value
				StoredProc.ParamByName("sOKTMO").value			= OKTMO
				StoredProc.ParamByName("sINN_CITIZENSHIP").value= bAgnQuery.FieldByname("INN_CITIZENSHIP").value
				StoredProc.ExecProc

				Query.SQL.Text = "select * from AGNBANKS where RN='"&agnbank_rn&"'"
				Query.Open
				'������� ������ �� �������� ������� "���������� ����������"
				StoredProc.StoredProcName="P_AGNBANKS_UPDATE"
				StoredProc.ParamByName("nCOMPANY").value		= 42903
				StoredProc.ParamByName("nRN").value				= agnbank_rn
				StoredProc.ParamByName("sBANKFCODEACC").value	= nodeNode.selectSingleNode("���").text
				StoredProc.ParamByName("sBANKACC").value		= nodeNode.selectSingleNode("��������").text
				StoredProc.ParamByName("sCODE").value			= "����_"&nodeNode.selectSingleNode("���").text
				StoredProc.ParamByName("sSWIFT").value			= Query.FieldByname("SWIFT").value
				StoredProc.ParamByName("sMEMBER_CODE").value	= Query.FieldByname("MEMBER_CODE").value
				StoredProc.ParamByName("sMEMBER_NAME").value	= Query.FieldByname("MEMBER_NAME").value
				StoredProc.ParamByName("sMEMBER_REG").value		= Query.FieldByname("MEMBER_REG").value
				StoredProc.ExecProc
			end if
		next
		MyFile.Write("INFO "&now()&vbTab&" ����� ���������. ����� ����������: "& object_counter & vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" � ������� �������� �� ������� ���������� � ������."&vbNewLine)
	end if

	'������� ���� "�����"
	Set Accounts = xmlParser.selectNodes("//���������������/������")
	If Accounts.length > 0 then
		'���������� ������ ������
		object_counter=0
		For Each nodeNode In Accounts
			account_rn		= NULL
			AccQueryIsEmpty	= NULL
			agnlist_rn		= NULL
			agnlist_name	= NULL
			agnbank_mnemo	= NULL
			lastcode		= NULL
			counter			= NULL
			newRN			= NULL
			
			object_counter=object_counter+1
					
			'������� RN ����������� �� ���� SAP
			Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&RTrim(nodeNode.selectSingleNode("��������").text)&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - ��� �������� "��� SAP"
			Query.Open
			agnlist_rn=Query.FieldByname("UNIT_RN").value
			Query.Close
			
			If not agnlist_rn=0 then
				If not nodeNode.selectSingleNode("����").text=""  then
					'������� ������ �����
					Query.SQL.Text = "select SNAME, SBANKACC, SCODE from v_AGNBANKS where SCODE='����_"&nodeNode.selectSingleNode("����").text&"'"
					Query.Open
					agnbank_mnemo = Query.FieldByname("SCODE").value
				else
					agnbank_mnemo = NULL
				end if
				
				
				'������� ������������ ����������� ��� ������������ �����
				Query.SQL.Text= "select AGNNAME from AGNLIST where RN='"&agnlist_rn&"'"
				Query.Open
				agnlist_name = Query.FieldByname("AGNNAME").value
				Query.Close
				
				Query.SQL.Text="select * from v_agnacc where AGNACC='"&nodeNode.selectSingleNode("����������").text&"' and AGNRN='"&agnlist_rn&"' and SBANKFCODEACC='"&agnbank_mnemo&"' order by STRCODE"
				Query.Open
				Query.Last
				
								
				If Query.IsEmpty then 
					MyFile.Write("INFO "&now()&vbTab&" � ����� �� ������ ���������� ���� � ������� "&nodeNode.selectSingleNode("����������").text&"  - ������ ����� ����."&vbNewLine)
					
					'������ �������� "��� ������"
					StoredProc.StoredProcName="FIND_AGNACC_LASTCODE"
					StoredProc.ParamByName("COMPANY").value	= 42903                                        '��� �������������
					StoredProc.ParamByName("AGNRN").value	= agnlist_rn                                        '��� �����������
					StoredProc.ExecProc
					lastcode = CInt(StoredProc.ParamByName("STRCODE").value)
					lastcode=lastcode+1
					lastcode=CStr(lastcode)
					counter = 4-len(lastcode)
					do while counter > 0
							lastcode="0"&lastcode
							counter=counter-1
					loop

					'������� ������ � ����� ���������� �����
					StoredProc.StoredProcName="P_AGNACC_INSERT"
					StoredProc.ParamByName("nCOMPANY").value		= 42903                                        '��� �������������
					StoredProc.ParamByName("nPRN").value			= agnlist_rn                                        '��� �����������
					StoredProc.ParamByName("sSTRCODE").value		= lastcode
					StoredProc.ParamByName("SAGNACC").value			= nodeNode.selectSingleNode("����������").text
					StoredProc.ParamByName("sAGNNAMEACC").value		= agnlist_name
					StoredProc.ParamByName("SAGNBANKS").value		= agnbank_mnemo
					StoredProc.ParamByName("NACCESS_FLAG").value	= 1
					StoredProc.ExecProc
					newRN = StoredProc.ParamByName("nRN").value

					REM '������ ��� 1� � ��������
					REM 'StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         '���������� ��� ������ ��� ��������
					REM 'StoredProc.ParamByName("PROPERTY").value="�������1�"
					REM 'StoredProc.ParamByName("UNITCODE").value="ContragentsBankAttrs"
					REM 'StoredProc.ParamByName("RN_SOTR").value=newRN
					REM 'StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("������").text)
					REM 'StoredProc.ParamByName("NUM_VAL").value=NULL
					REM 'StoredProc.ExecProc
					Query.Close
				else
					MyFile.Write("INFO "&now()&vbTab&" � ����� ������ ���������� ���� � ������� "&nodeNode.selectSingleNode("����������").text&"  - �������� ������������ ����."&vbNewLine)
					
					account_rn = Query.FieldByname("RN").value

					'������ ����� ������ ��� ������� ������ � ���������� �����
					Query.SQL.Text = "select * from AGNACC where RN = '"&account_rn&"'"
					Query.Open
					
					STRCODE			= Query.FieldByname("STRCODE").value
					BANKNAMEACC		= Query.FieldByname("BANKNAMEACC").value
					BANKFCODEACC	= Query.FieldByname("BANKFCODEACC").value
					BANKACC			= Query.FieldByname("BANKACC").value
					BANKCITYACC		= Query.FieldByname("BANKCITYACC").value
					OPEN_DATE		= Query.FieldByname("OPEN_DATE").value
					CLOSE_DATE		= Query.FieldByname("CLOSE_DATE").value
					COUNTRY_CODE	= Query.FieldByname("COUNTRY_CODE").value
					SWIFT			= Query.FieldByname("SWIFT").value
					REGION			= Query.FieldByname("REGION").value
					DISTRICT		= Query.FieldByname("DISTRICT").value
					BANKACC_TYPE	= Query.FieldByname("BANKACC_TYPE").value
					sCURRENCY		= Query.FieldByname("CURRENCY").value
					CORR_AGNACC		= Query.FieldByname("CORR_AGNACC").value
					CARDNUMB		= Query.FieldByname("CARDNUMB").value
					AGNTREAS		= Query.FieldByname("AGNTREAS").value
					REAS_AGNACC		= Query.FieldByname("TREAS_AGNACC").value
					INTERMEDIARY	= Query.FieldByname("INTERMEDIARY").value
					INTERMED_ACC	= Query.FieldByname("INTERMED_ACC").value
					
					'��������� � NULL ��� �������� 0
					If BANKACC_TYPE = 0 then
						BANKACC_TYPE = NULL
					else
						Query.SQL.Text = "select CODE from BANKACCTYPES where RN='"&BANKACC_TYPE&"'"
						Query.Open
						BANKACC_TYPE = Query.FieldByname("CODE").value
					end if	
					If sCURRENCY = 0 then
						sCURRENCY = NULL
					else
						Query.SQL.Text = "select CURCODE from CURNAMES where RN='"&sCURRENCY&"'"
						Query.Open
						sCURRENCY = Query.FieldByname("CURCODE").value
					end if	
					
					CORR_AGNACC = NULL
					
					If AGNTREAS = 0 then
						AGNTREAS = NULL
					end if	
					If REAS_AGNACC = 0 then
						REAS_AGNACC = NULL
					end if	
					If INTERMEDIARY = 0 then
						INTERMEDIARY = NULL
					end if	
					If INTERMED_ACC = 0 then
						INTERMED_ACC = NULL
					end if	

					'������� ������ � ���������� �����
					StoredProc.StoredProcName="P_AGNACC_UPDATE"
					StoredProc.ParamByName("nCOMPANY").value		= 42903
					StoredProc.ParamByName("nRN").value				= account_rn
					StoredProc.ParamByName("sSTRCODE").value		= STRCODE
					StoredProc.ParamByName("SAGNACC").value			= nodeNode.selectSingleNode("����������").text
					StoredProc.ParamByName("sAGNNAMEACC").value		= agnlist_name
					StoredProc.ParamByName("sBANKNAMEACC").value	= BANKNAMEACC
					StoredProc.ParamByName("sBANKFCODEACC").value	= BANKFCODEACC
					StoredProc.ParamByName("sBANKACC").value		= BANKACC
					StoredProc.ParamByName("sBANKCITYACC").value	= BANKCITYACC
					StoredProc.ParamByName("sAGNBANKS").value		= agnbank_mnemo
					StoredProc.ParamByName("dOPEN_DATE").value		= OPEN_DATE
					StoredProc.ParamByName("dCLOSE_DATE").value		= CLOSE_DATE
					StoredProc.ParamByName("sCOUNTRY_CODE").value	= COUNTRY_CODE
					StoredProc.ParamByName("nACCESS_FLAG").value	= 1
					StoredProc.ParamByName("sSWIFT").value			= SWIFT
					StoredProc.ParamByName("sREGION").value			= REGION
					StoredProc.ParamByName("sDISTRICT").value		= DISTRICT
					StoredProc.ParamByName("sBANKACC_TYPE").value	= BANKACC_TYPE
					StoredProc.ParamByName("sCURRENCY").value		= sCURRENCY
					StoredProc.ParamByName("sCORR_AGNACC").value	= CORR_AGNACC
					StoredProc.ParamByName("sCARDNUMB").value		= CARDNUMB
					StoredProc.ParamByName("sAGNTREAS").value		= AGNTREAS
					StoredProc.ParamByName("sTREAS_AGNACC").value	= REAS_AGNACC
					StoredProc.ParamByName("sINTERMEDIARY").value	= INTERMEDIARY
					StoredProc.ParamByName("sINTERMED_ACC").value	= INTERMED_ACC
					StoredProc.ExecProc
				end if
			else
				If nodeNode.selectSingleNode("������������������").text = "true" or nodeNode.selectSingleNode("����������").text ="������������ ��� ���"  then
					MyFile.Write("INFO "&now()&vbTab&" C��� � ������� "&nodeNode.selectSingleNode("����������").text&" ����������� ��� <������������ ���> - ��������� ����."&vbNewLine)
				elseIf nodeNode.selectSingleNode("��������").text = "" then
					MyFile.Write("INFO "&now()&vbTab&"� ����� � ������� "&nodeNode.selectSingleNode("����������").text&" �� ������ �������� - ��������� ����."&vbNewLine)
				elseIf nodeNode.selectSingleNode("��������").text = "����������" then
					MyFile.Write("INFO "&now()&vbTab&" C��� � ������� "&nodeNode.selectSingleNode("����������").text&" �������� ���������� - ��������� ����."&vbNewLine)
				else
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" � ����� �� ������ ���������� � ����� SAP "&nodeNode.selectSingleNode("��������").text&" �������� ����������� ���� � ������� "&nodeNode.selectSingleNode("����������").text&"  - �� ���� �������/�������� ����."&vbNewLine)
				end if
			end if
		next
		MyFile.Write("INFO "&now()&vbTab&" ���������� ����� ���������. ����� ����������: "& object_counter & vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" � ������� �������� �� ������� ���������� � ���������� ������."&vbNewLine)
	end if

	'������� ���� "��������"'
	Set Contracts = xmlParser.selectNodes("//��������������������/������")
	If Contracts.length > 0 then
		object_counter = 0
		For Each nodeNode In Contracts
			comment					= NULL
			comment_array			= NULL
			contract_complex_number	= NULL
			contract_number_array	= NULL
			doc_type_RN				= NULL
			old_RN					= NULL
			agn_abbr				= NULL
			agn_rn					= NULL
			orgaccbik				= NULL
			orgacc					= NULL
			jur_strcode				= NULL
			agnaccbik				= NULL
			agnacc					= NULL
			agn_strcode				= NULL
			INOUT_SIGN				= NULL
			ext_agreement			= 0
			PRN						= NULL
			executive				= NULL
			subdiv					= NULL
			scurrency				= NULL
			newContract				= NULL
			doc_numb				= NULL
			warning					= NULL
			
			object_counter = object_counter+1
			
			Query.SQL.Text = "select AGNABBR from agnlist where agnname like upper('%"&nodeNode.selectSingleNode("������������������������").text&"%') and EMP=1 order by RN DESC"
			Query.Open
			If Query.IsEmpty and not InStr(nodeNode.selectSingleNode("�����������").text, "�������")=0 then
				MyFile.Write("INFO "&now()&vbTab&" ������� "&comment&" ("&trim(nodeNode.selectSingleNode("������").text)&") - ����������, ��������� �������."&vbNewLine)
			else
				'������ ������� �� ������ 1�
				Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value like '%"&trim(nodeNode.selectSingleNode("������").text)&"%' and docs_prop_rn='104582667' and unitcode='Contracts'"        ' 87456099 - ��� �������� "��� 1�"
				Query.Open 
				If Query.IsEmpty then
					'���������� �������� �� ������, ���� �� ������� ��������� �����������
					doc_type		= "0000"
					doc_pref		= date&"/"&timer
					'������� ��������� ���������� �����
					StoredProc.StoredProcName="P_CONTRACTS_GETNEXTNUMB"         
					StoredProc.ParamByName("NCOMPANY").value=42903
					StoredProc.ParamByName("SJUR_PERS").value="�� ���"
					StoredProc.ParamByName("DDOC_DATE").value=ConvDate(nodeNode.selectSingleNode("����").text)
					StoredProc.ParamByName("SDOC_TYPE").value=doc_type
					StoredProc.ParamByName("SDOC_PREF").value=doc_pref
					StoredProc.ExecProc
					doc_numb		= StoredProc.ParamByName("SDOC_NUMB").value

					'���� ����������� �������� - ��������� �������� �� ���� ������ �� ����� ������������ ��������
					comment = nodeNode.selectSingleNode("�����������").text
					If not len(comment)=0 then
						'�������� ����������� �� ��������� �����: ���, ������� � �����, �������� ���������� ��������� ������
						comment_array			= Split(LTrim(comment), ",")
						if UBound(comment_array)>0 then
							contract_complex_number	= RTrim(LTrim(comment_array(1)))
							contract_number_array	= Split(contract_complex_number, "-")
							if UBound(contract_number_array)>0 then
								doc_type		= comment_array(0)
								doc_pref		= contract_number_array(0)
								doc_numb		= contract_number_array(1)
																
								'������ ��� ���� ���������
								Query.SQL.Text = "select RN from DOCTYPES where DOCCODE='"&doc_type&"'"
								Query.Open
								doc_type_RN	= Query.FieldByname("RN").value
								Query.Close
								
								
								'������ ������� �� ����, �������� � ������
								Query.SQL.Text = "select RN from contracts where doc_type='"&doc_type_RN&"' and DOC_PREF like '%"&doc_pref&"' and DOC_NUMB like '%"&doc_numb&"'"
								Query.Open
								If not Query.IsEmpty and (nodeNode.selectSingleNode("��������������").text="00000000-0000-0000-0000-000000000000" or len(nodeNode.selectSingleNode("��������������").text)=0) then
									newContract = False
									old_RN = Query.FieldByname("RN").value		'�������� ��� ��� ������������� ��������
								else
									newContract = True
								end if
								Query.Close
							else
								newContract = True
							end if
						else
							newContract = True
						end if
					else
						newContract = True
					end If				
				else
					newContract = False
					old_RN = Query.FieldByname("UNIT_RN").value		'�������� ��� ��� ������������� ��������
				end if
				
				'���� ������� �� ������ �� ���� 1� ��� ������ ����������� - �������� �����
				If newContract then
					MyFile.Write("INFO "&now()&vbTab&" � ����� �� ������ ������� "&comment&" ("&trim(nodeNode.selectSingleNode("������").text)&"), ���� �� ������� ��������� ��������� ���� ����������� - ������ ������� ����� �������."&vbNewLine)		
					
					'������� �������� ����������� �� ���� SAP
					Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&RTrim(nodeNode.selectSingleNode("��������").text)&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - ��� �������� "��� SAP" (1�������� ����� �������� � ���������)
					Query.Open
					Query.SQL.Text = "select AGNABBR, RN, AGNNAME from AGNLIST where RN='" & Query.FieldByname("UNIT_RN").value & "'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value
					agn_rn = Query.FieldByname("RN").value
					agn_name = Query.FieldByname("AGNNAME").value
					Query.Close
					
					'������� ����� ������ ����������� ����� ����������� ����� ����� ����� �����
					If not nodeNode.selectSingleNode("���������������").text = "" then
						orgaccbik		= Trim(nodeNode.selectSingleNode("������������������").text)
						orgacc			= Trim(nodeNode.selectSingleNode("���������������").text)
						Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & orgaccbik & "'"
						Query.Open
						If not Query.IsEmpty then
							Query.SQL.Text	= "select STRCODE from AGNACC where AGNACC='"& orgacc &"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='5805775'"
							Query.Open
							jur_strcode		= Query.FieldByname("STRCODE").value
							Query.Close
						else
							MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ��� ����������� ����� � ������� "&nodeNode.selectSingleNode("���������������").text&" � ����� �� ������ ��� "&nodeNode.selectSingleNode("������������������").text&" - ����� ������� ��� �������."&vbNewLine)
						end if
					else
						jur_strcode		= NULL
					end if
					
					'������� ����� ������ ����������� ����� ����������� ����� ����� ����� �����
					if not nodeNode.selectSingleNode("���������������").text="" then
						agnaccbik		= Trim(nodeNode.selectSingleNode("������������������").text)
						agnacc			= Trim(nodeNode.selectSingleNode("���������������").text)
						Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & agnaccbik & "'"
						Query.Open
						If not Query.IsEmpty then
							Query.SQL.Text = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='"&agn_rn&"'"
							Query.Open
							If not Query.IsEmpty then
								agn_strcode = Query.FieldByname("STRCODE").value
								note = nodeNode.selectSingleNode("���������������").text
							else
								MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ���������� ���� � ������� "&nodeNode.selectSingleNode("���������������").text&" �� ����������� ����������� "&agn_abbr&"  - � �������� <<"&comment&">> ����� ������ ��������� ���� �����������"&vbNewLine)
								Query.SQL.Text = "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
								Query.Open
								agn_strcode = Query.FieldByname("STRCODE").value
								note = "� �������� ������ ��������� ���������� ���� ����������� - �������� ���������� ����!!!"
							end if
							Query.Close
						else
							MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ��� ����������� ����� � ������� "&nodeNode.selectSingleNode("���������������").text&" � ����� �� ������ ��� "&nodeNode.selectSingleNode("������������������").text&" - ����� ������� ��� �������."&vbNewLine)
						end if					
					else
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" � ����� �� ������ ���������� ���� ����������� "&agn_abbr&" - � �������� <<"&comment&">> ����� ������ ��������� ���� �����������"&vbNewLine)
						Query.SQL.Text = "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
						Query.Open
						if not Query.IsEmpty then
							agn_strcode = Query.FieldByname("STRCODE").value
						else
							'������� ������ � ����� ���������� �����
							StoredProc.StoredProcName="P_AGNACC_INSERT"
							StoredProc.ParamByName("nCOMPANY").value		= 42903 
							StoredProc.ParamByName("nPRN").value			= agn_rn
							StoredProc.ParamByName("sSTRCODE").value		= "0001"
							StoredProc.ParamByName("SAGNACC").value			= NULL 'nodeNode.selectSingleNode("����������").text
							StoredProc.ParamByName("sAGNNAMEACC").value		= agn_name
							StoredProc.ParamByName("SAGNBANKS").value		= NULL
							StoredProc.ParamByName("NACCESS_FLAG").value	= 1
							StoredProc.ExecProc
							agn_strcode = "0001"
						end if
						note = "� �������� ������ ��������� ���������� ���� ����������� - �������� ���������� ����!!!"
					end if
					
					'�������� ��������� �����
					if len(nodeNode.selectSingleNode("������������").text)>0 then
						INOUT_SIGN = 0
					else
						INOUT_SIGN = 1
					end if				
									
					'������� ��� ������������� �� �������: "��� ����������� -> ��� ����������� -> ��� ���������� -> ������ � ������� ��������� -> ��� ������������� -> ��� �������������"
					if nodeNode.selectSingleNode("������������������������").text="������ ������ ���������" then
						executive = "0001 ������ �.�."		'������ �.�. - �������, �� ��� ������ ����������� ��� - �������� ������������� "����� ������"
						subdiv = "�����.13.15"
					else
						Query.SQL.Text = "select a.AGNABBR, a.AGNNAME, a.RN, b.code from agnlist a, CLNPERSONS b where a.rn=b.pers_agent and agnname like upper('%"&nodeNode.selectSingleNode("������������������������").text&"%') and not b.crn=2503442 and EMP=1 order by RN DESC"	'������ �� ����������� � ����������� �� ����� ���������
						Query.Open
						executive		= Query.FieldByname("AGNABBR").value				
						Query.SQL.Text = "select RN from CLNPERSONS where PERS_AGENT='"&Query.FieldByname("RN").value&"'"
						Query.Open								
						Query.SQL.Text = "select DEPTRN from CLNPSPFM where persrn='"&Query.FieldByname("RN").value&"' and endeng is null"
						Query.Open				
						Query.SQL.Text = "select CODE from INS_DEPARTMENT where rn='"&Query.FieldByname("DEPTRN").value&"'"
						Query.Open
						subdiv = Query.FieldByname("CODE").value
						Query.Close
					end if
					
					'������� ������������ ������ �� �� ����
					Query.SQL.Text = "select INTCODE from curnames where curcode='"&nodeNode.selectSingleNode("��������������������").text&"'"
					Query.Open
					If Query.IsEmpty then
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ������ � ����� "&nodeNode.selectSingleNode("��������������������").text&" �� �������� <<"&comment&">> �� �������. ������������ �������� ��-��������� - RUR"&vbNewLine)
						scurrency = "RUR"
					else
						scurrency = nodeNode.selectSingleNode("��������������������").text
					end if
					Query.Close
					
					If nodeNode.selectSingleNode("��������������").text="0001-01-01" then
						endDate = "01.01.0001"	'NULL
					Else
						endDate = ConvDate(nodeNode.selectSingleNode("��������������").text)
					end if
					
					If agn_abbr="" then
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" � ����� �� ������ ���������� � ����� "&nodeNode.selectSingleNode("��������").text&" �� �������� "&comment&" - ��������� �������."&vbNewLine)
					else
					
						'������� ��� ����� � ������ ��� �����
						If nodeNode.selectSingleNode("�����������").text = "� �����������" then
							sumType = 1
							acc_kind = 0
						elseif nodeNode.selectSingleNode("�����������").text = "� �����������" then
							sumType = 2
							acc_kind = 1
						else
							sumType = 0
							acc_kind = 0
						end if
						
						'��������� - ������������� ��� ����� �������
						if nodeNode.selectSingleNode("��������������").text="00000000-0000-0000-0000-000000000000" or len(nodeNode.selectSingleNode("��������������").text)=0 then
							'������� ������ � ����� ��������
							StoredProc.StoredProcName="P_CONTRACTS_INSERT"
							StoredProc.ParamByName("nCOMPANY").value		= 42903
							StoredProc.ParamByName("nCRN").value			= 104583519                                        '��� ��������
							StoredProc.ParamByName("nPRN").value			= PRN                                        '��� ��������� ��������
							StoredProc.ParamByName("SJUR_PERS").value		= "�� ���"
							StoredProc.ParamByName("SJUR_ACC").value		= jur_strcode
							StoredProc.ParamByName("SDOC_TYPE").value		= doc_type
							StoredProc.ParamByName("SDOC_PREF").value		= doc_pref
							StoredProc.ParamByName("SDOC_NUMB").value		= doc_numb
							StoredProc.ParamByName("DDOC_DATE").value		= ConvDate(nodeNode.selectSingleNode("����").text)
							StoredProc.ParamByName("SEXT_NUMBER").value		= nodeNode.selectSingleNode("������������").text
							StoredProc.ParamByName("NINOUT_SIGN").value		= INOUT_SIGN	'������ ��������, ������=0, ����=1
							StoredProc.ParamByName("NFALSE_DOC").value		= 0				'������ ��������, ����=0, ������=1
							StoredProc.ParamByName("NEXT_AGREEMENT").value	= ext_agreement	'������ �������������, ����=0, ������=1
							StoredProc.ParamByName("SAGENT").value			= agn_abbr
							StoredProc.ParamByName("SAGNACC").value			= agn_strcode
							StoredProc.ParamByName("SEXECUTIVE").value		= executive
							StoredProc.ParamByName("SSUBDIVISION").value	= subdiv
							StoredProc.ParamByName("DBEGIN_DATE").value		= ConvDate(nodeNode.selectSingleNode("�������������").text)
							StoredProc.ParamByName("DEND_DATE").value		= endDate
							StoredProc.ParamByName("NSUM_TYPE").value		= 1
							StoredProc.ParamByName("NDOC_SUM").value		= 0
							StoredProc.ParamByName("NDOC_SUMTAX").value		= Replace(nodeNode.selectSingleNode("�������������").text, ".", ",")
							StoredProc.ParamByName("NDOC_SUM_NDS").value	= 0
							StoredProc.ParamByName("NAUTOCALC_SIGN").value	= 1
							StoredProc.ParamByName("SCURRENCY").value		= scurrency
							StoredProc.ParamByName("NCURCOURS").value		= 1
							StoredProc.ParamByName("NCURBASE").value		= 1
							StoredProc.ParamByName("SSUBJECT").value		= nodeNode.selectSingleNode("���������������").text
							StoredProc.ParamByName("SNOTE").value			= note
							StoredProc.ExecProc
							newRN = StoredProc.ParamByName("nRN").value
							
							'������� ������ � �������� ���������
							StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         '������1�
							StoredProc.ParamByName("PROPERTY").value="������1�"
							StoredProc.ParamByName("UNITCODE").value="Contracts"
							StoredProc.ParamByName("RN_SOTR").value=newRN
							StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("������").text)
							StoredProc.ParamByName("NUM_VAL").value=NULL
							StoredProc.ExecProc
							
							'������� ������ ���� � ��������
							StoredProc.StoredProcName="P_STAGES_INSERT"        
							StoredProc.ParamByName("nCOMPANY").value		= 42903
							StoredProc.ParamByName("nPRN").value			= newRN
							StoredProc.ParamByName("SNUMB").value			= "1"
							StoredProc.ParamByName("NEXT_AGREEMENT").value	= 0
							StoredProc.ParamByName("NSIGN_SUM").value		= 1
							StoredProc.ParamByName("DBEGIN_DATE").value		= ConvDate(nodeNode.selectSingleNode("�������������").text)
							StoredProc.ParamByName("DEND_DATE").value		= endDate
							StoredProc.ParamByName("SJUR_ACC").value		= jur_strcode
							StoredProc.ParamByName("NSUM_TYPE").value		= sumType				'������ �����
							StoredProc.ParamByName("NSTAGE_SUM").value		= 0
							StoredProc.ParamByName("NSTAGE_SUMTAX").value	= Replace(nodeNode.selectSingleNode("�������������").text, ".", ",")
							StoredProc.ParamByName("NSTAGE_SUM_NDS").value	= 0
							StoredProc.ParamByName("NAUTOCALC_SIGN").value	= 1
							StoredProc.ParamByName("SDESCRIPTION").value	= nodeNode.selectSingleNode("���������������").text
							StoredProc.ParamByName("SCOMMENTS").value		= nodeNode.selectSingleNode("���������������").text
							StoredProc.ParamByName("NFACEACC_EXIST").value	= 0
							StoredProc.ParamByName("SFACEACCCRN").value		= "test" 'GetFaceAccCat(subdiv)'������� �� �������������
							StoredProc.ParamByName("SAGENT").value			= agn_abbr
							StoredProc.ParamByName("SFACEACC").value		= doc_pref&"/"&doc_numb&"/1"
							StoredProc.ParamByName("NACC_KIND").value		= acc_kind
							StoredProc.ParamByName("SEXECUTIVE").value		= executive
							StoredProc.ParamByName("SCURRENCY").value		= scurrency
							StoredProc.ParamByName("NCREDIT_SUM").value		= 0
							StoredProc.ParamByName("SAGNACC").value			= agn_strcode
							StoredProc.ParamByName("SSUBDIV").value			= subdiv
							StoredProc.ParamByName("NDISCOUNT").value		= 0
							StoredProc.ParamByName("NPRICE_TYPE").value		= 0
							StoredProc.ParamByName("NSIGNTAX").value		= 1
							StoredProc.ParamByName("NSAME_NOMN").value		= 0					
							StoredProc.ExecProc
							
							Set ExtraData = nodeNode.selectNodes("������������/������")
							If ExtraData.length > 0 then
								For Each node In ExtraData
									AttribName = node.selectSingleNode("���").text
									If AttribName = "������ �������" or AttribName = "���������� �������� ������" or AttribName = "���������� ������������ ������" or AttribName = "��������� ���� �������" or AttribName = "� ���������� ���������" then
										Query.SQL.text = "select CODE from docs_props where name = '"&AttribName&"'"
										Query.Open	
										
										StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         
										StoredProc.ParamByName("PROPERTY").value=Query.FieldByname("CODE").value
										StoredProc.ParamByName("UNITCODE").value="Contracts"
										StoredProc.ParamByName("RN_SOTR").value=newRN
										StoredProc.ParamByName("ST_VAL").value=node.selectSingleNode("��������").text
										StoredProc.ParamByName("NUM_VAL").value=NULL
										StoredProc.ExecProc
									elseif AttribName = "�������������� � ��������� ���/��. �������" then
										If node.selectSingleNode("��������").text="��" then
											val = 1
										else
											val = 0
										end if
										StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"        
										StoredProc.ParamByName("PROPERTY").value="������� ���/�� ����."
										StoredProc.ParamByName("UNITCODE").value="Contracts"
										StoredProc.ParamByName("RN_SOTR").value=newRN
										StoredProc.ParamByName("ST_VAL").value=NULL
										StoredProc.ParamByName("NUM_VAL").value=val
										StoredProc.ExecProc
									end if
								next
							end if
						else
							Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&nodeNode.selectSingleNode("��������������").text&"' and docs_prop_rn='104582667' and unitcode='Contracts'"        ' 87456099 - ��� �������� "��� 1�"
							Query.Open
							UNIT_RN = Query.FieldByname("UNIT_RN").value
							If Query.IsEmpty then
								MyFile.Write(vbTab&"ERROR "&now()&vbTab&" � ����� �� ������ ������������ ������� � ����� 1� "&nodeNode.selectSingleNode("��������������").text&" - �� ���� ������� ����� ����"&vbNewLine)
							else	
								Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&nodeNode.selectSingleNode("������").text&"' and docs_prop_rn='109199304' and unitcode='ContractsStages'"        ' 109199304 - ��� �������� "��� 1�"
								Query.Open
								If Query.IsEmpty then
									MyFile.Write(vbTab&"INFO "&now()&vbTab&" � ����� ������ ������������ ������� � ����� 1� "&nodeNode.selectSingleNode("��������������").text&" - ������ ����� ����"&vbNewLine)
									
									'������� ������� ��������� �����
									Query.SQL.Text="select DOC_PREF, DOC_NUMB from CONTRACTS where RN='"&UNIT_RN&"'"
									Query.Open
									doc_pref = Trim(Query.FieldByname("DOC_PREF").value)
									doc_numb = Trim(Query.FieldByname("DOC_NUMB").value)
																	
									'������� ���������� �����
									StoredProc.StoredProcName="P_STAGES_GETNEXTNUMB"
									StoredProc.ParamByName("NCOMPANY").value		= 42903
									StoredProc.ParamByName("NPRN").value			= UNIT_RN
									StoredProc.ExecProc
									snumb = StoredProc.ParamByName("SNUMB_MAX").value
									
									'������� ����� ���� � ��������
									StoredProc.StoredProcName="P_STAGES_INSERT"        
									StoredProc.ParamByName("nCOMPANY").value		= 42903
									StoredProc.ParamByName("nPRN").value			= UNIT_RN
									StoredProc.ParamByName("SNUMB").value			= snumb
									StoredProc.ParamByName("NEXT_AGREEMENT").value	= 1
									StoredProc.ParamByName("NSIGN_SUM").value		= 1
									StoredProc.ParamByName("DBEGIN_DATE").value		= ConvDate(nodeNode.selectSingleNode("�������������").text)
									StoredProc.ParamByName("DEND_DATE").value		= endDate
									StoredProc.ParamByName("SJUR_ACC").value		= jur_strcode
									StoredProc.ParamByName("NSUM_TYPE").value		= sumType				'������ �����
									StoredProc.ParamByName("NSTAGE_SUM").value		= 0
									StoredProc.ParamByName("NSTAGE_SUMTAX").value	= Replace(nodeNode.selectSingleNode("�������������").text, ".", ",")
									StoredProc.ParamByName("NSTAGE_SUM_NDS").value	= 0
									StoredProc.ParamByName("NAUTOCALC_SIGN").value	= 1
									StoredProc.ParamByName("SDESCRIPTION").value	= nodeNode.selectSingleNode("���������������").text
									StoredProc.ParamByName("SCOMMENTS").value		= nodeNode.selectSingleNode("���������������").text
									StoredProc.ParamByName("NFACEACC_EXIST").value	= 0
									StoredProc.ParamByName("SFACEACCCRN").value		= "test" 'GetFaceAccCat(subdiv)'������� �� �������������
									StoredProc.ParamByName("SAGENT").value			= agn_abbr
									StoredProc.ParamByName("SFACEACC").value		= doc_pref&"/"&doc_numb&"/"&snumb
									StoredProc.ParamByName("NACC_KIND").value		= 0
									StoredProc.ParamByName("SEXECUTIVE").value		= executive
									StoredProc.ParamByName("SCURRENCY").value		= scurrency
									StoredProc.ParamByName("NCREDIT_SUM").value		= 0
									StoredProc.ParamByName("SAGNACC").value			= agn_strcode
									StoredProc.ParamByName("SSUBDIV").value			= subdiv
									StoredProc.ParamByName("NDISCOUNT").value		= 0
									StoredProc.ParamByName("NPRICE_TYPE").value		= 0
									StoredProc.ParamByName("NSIGNTAX").value		= 1
									StoredProc.ParamByName("NSAME_NOMN").value		= 0					
									StoredProc.ExecProc
									newRN = StoredProc.ParamByName("nRN").value
									
									'������� ������ � �������� ���������
									StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         '������1�
									StoredProc.ParamByName("PROPERTY").value="����������1�"
									StoredProc.ParamByName("UNITCODE").value="ContractsStages"
									StoredProc.ParamByName("RN_SOTR").value=newRN
									StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("������").text)
									StoredProc.ParamByName("NUM_VAL").value=NULL
									StoredProc.ExecProc
								else
									MyFile.Write("INFO "&now()&vbTab&" � ����� ������ ���� "&comment&" ("&trim(nodeNode.selectSingleNode("������").text)&") - ��������� ����."&vbNewLine)
								end if
							end if
						end if
					end if
				else
					MyFile.Write("INFO "&now()&vbTab&" � ����� ������ ������� "&comment&" ("&trim(nodeNode.selectSingleNode("������").text)&") - ��������� �������."&vbNewLine)
					'������� ������ � �������� ���������
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         '������1�
					StoredProc.ParamByName("PROPERTY").value="������1�"
					StoredProc.ParamByName("UNITCODE").value="Contracts"
					StoredProc.ParamByName("RN_SOTR").value=old_RN
					StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("������").text)
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc
				end if
			end if
		next
		MyFile.Write("INFO "&now()&vbTab&" �������� ���������. ����� ����������: "& object_counter &vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" � ������� �������� �� ������� ���������� � ���������."&vbNewLine&vbNewLine)
	end if
	
	'������� ���� "������������������������"'
	Set Outcoming = xmlParser.selectNodes("//������������������������/������")
	If Outcoming.length > 0 then
		object_counter = 0
		For Each nodeNode In Outcoming
			DATE_array		= NULL
			DATE_year		= NULL
			banknumb		= NULL
			DATETIME_array	= NULL
			DATEONLY		= NULL
			docpref 		= NULL
			docnumb 		= NULL
			doctype 		= NULL
			delimiter		= NULL
			docdate 		= NULL
			jur_strcode		= NULL
			agn_abbr		= NULL
			agnaccbik		= NULL
			agnacc			= NULL
			agn_strcode		= NULL
			typeoper_mnemo	= NULL
			scurrency		= NULL
			newRN			= NULL
			agn_rn			= NULL
			SPAY_NOTE		= NULL
						
			object_counter = object_counter + 1
			
			Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value like '%"&trim(nodeNode.selectSingleNode("������").text)&"%' and docs_prop_rn='104582883' and unitcode='BankDocuments'"
			Query.Open
			If Query.IsEmpty then
				MyFile.Write("INFO "&now()&vbTab&" � ����� �� ������ �������� �������� �� � ����� "&trim(nodeNode.selectSingleNode("������").text)&"  "&nodeNode.selectSingleNode("�������").text&" "&nodeNode.selectSingleNode("������").text&" - ������ ����� ��������."&vbNewLine)
				
				'������� ��� �� ����
				DATE_array	= Split(nodeNode.selectSingleNode("������").text, "-")
				DATE_year	= DATE_array(0)
				
				'������� ���������� �����
				StoredProc.StoredProcName="P_BANKDOCS_GETNEXTNUMB"
				StoredProc.ParamByName("NCOMPANY").value		= 42903
				StoredProc.ParamByName("SJUR_PERS").value		= "�� ���"
				StoredProc.ParamByName("DBANK_DOCDATE").value	= ConvDate(nodeNode.selectSingleNode("������").text)
				StoredProc.ParamByName("SBANK_DOCTYPE").value	= "������/�"
				StoredProc.ParamByName("SBANK_DOCPREF").value	= DATE_year
				StoredProc.ExecProc
				banknumb = StoredProc.ParamByName("SBANK_NUMB").value
				
				'������� ���� �� ������ ���� ����������
				DATETIME_array	= Split(nodeNode.selectSingleNode("����").text, "T")
				DATEONLY = DATETIME_array(0)
				
				'������� ������ ��������, ���� � ��������� ������ 1 �������
				Set TableRows = nodeNode.selectNodes("��������/������")
				If TableRows.length = 1 then
					Set Node = TableRows.nextNode()
					Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&Node.selectSingleNode("�������").text&"' and docs_prop_rn='104582667' and unitcode='Contracts'"        ' 87456099 - ��� �������� "��� 1�"
					Query.Open
					If not Query.IsEmpty then
						Query.SQL.Text="select DOC_TYPE, DOC_PREF, DOC_NUMB, DOC_DATE from CONTRACTS where RN='"&Query.FieldByname("UNIT_RN").value&"'"
						Query.Open
						docdate = Query.FieldByname("DOC_DATE").value
						docpref = LTrim(Query.FieldByname("DOC_PREF").value)
						docnumb = LTrim(Query.FieldByname("DOC_NUMB").value)
						delimiter = "-"
						Query.SQL.Text="select DOCCODE from DOCTYPES where RN='"&Query.FieldByname("DOC_TYPE").value&"'"
						Query.Open
						doctype = Query.FieldByname("DOCCODE").value
						Query.Close
					end if
					
					'������� ������ ��� � �� ��������
					If Node.selectSingleNode("���������").text="������" or Node.selectSingleNode("���������").text="" then
						TAXrate	= 0
						Tax		= 0
					else
						TAXrate	= Mid(Node.selectSingleNode("���������").text, 4)
						Tax		= Replace(Node.selectSingleNode("��������").text, ".", ",")
					End if
				end if				

				'������ �������� ����� ����������� �� �� ���� SAP
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - ��� �������� "��� SAP" (1�������� ����� �������� � ���������)
				Query.Open
				Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				SAGENT_FROM = Query.FieldByname("AGNABBR").value
				Query.Close
				
				'������� ����� ������ ����������� ����� ����������� ����� ����� ����� �����
				Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & Trim(nodeNode.selectSingleNode("������������������").text) & "'"
				Query.Open
				If not Query.IsEmpty then
					Query.SQL.Text	= "select STRCODE from AGNACC where AGNACC='"& Trim(nodeNode.selectSingleNode("���������������").text) &"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='5805775'"
					Query.Open
					jur_strcode		= Query.FieldByname("STRCODE").value
					Query.Close
				else
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ��� ����������� ����� � ������� "&nodeNode.selectSingleNode("���������������").text&" � ����� �� ������ ��� "&nodeNode.selectSingleNode("������������������").text&" - ����� ������� ��� �������."&vbNewLine)
				end if
				
				
				'���� ����������� � ���� � ����������� �� ��������� �������
				if nodeNode.selectSingleNode("�����������").text = "�������������������" then
					'���������� - ���
					Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
					Query.Open
					agn_rn = Query.FieldByname("UNIT_RN").value
					Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value
					
					'C��� - �� ��� ������
					agnaccbik		= Trim(nodeNode.selectSingleNode("������������������").text)
					agnacc			= Trim(nodeNode.selectSingleNode("���������������").text)
					If not nodeNode.selectSingleNode("������������������").text="" then
						Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & agnaccbik & "'"
						Query.Open
						bankrn = Query.FieldByname("NRN").value
					else
						bankrn = NULL
					end if
					If not IsNull(bankrn) then
						Query.SQL.Text 	= "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"&bankrn&"' and AGNRN='"&agn_rn&"'"
						Query.Open
						If not Query.IsEmpty then
							agn_strcode = Query.FieldByname("STRCODE").value
						else
							'�� ����� �� ������ - �� �������� �����							
							Query.SQL.Text = "select STRCODE from agnacc where agnrn='5805775' and agnbanks='"&bankrn&"'"
							Query.Open
							agn_strcode = Query.FieldByname("STRCODE").value
						end if
					else
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ��� ����������� ����� � ������� "&nodeNode.selectSingleNode("���������������").text&" � ����� �� ������ ��� "&nodeNode.selectSingleNode("������������������").text&" - ����� ������� ��� �������."&vbNewLine)
					end if
				else
					'���� �����������
					Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='"&RTrim(nodeNode.selectSingleNode("����������").text)&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
					Query.Open
					agn_rn = Query.FieldByname("UNIT_RN").value
					Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value
					
					'���� ����
					If nodeNode.selectSingleNode("���������������").text = "" then '������ ���������� ����
						Query.SQL.Text 	= "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
						Query.Open
						agn_strcode 	= Query.FieldByname("STRCODE").value
					else
						agnaccbik		= Trim(nodeNode.selectSingleNode("������������������").text)
						agnacc			= Trim(nodeNode.selectSingleNode("���������������").text)
						If not nodeNode.selectSingleNode("������������������").text="" then
							Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & agnaccbik & "'"
							Query.Open
							bankrn = Query.FieldByname("NRN").value
						else
							bankrn = NULL
						end if
						If IsNull(bankrn) then
							Query.SQL.Text 	= "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS is NULL and AGNRN='"&agn_rn&"'"
							Query.Open
							agn_strcode 	= Query.FieldByname("STRCODE").value
						else
							Query.SQL.Text 	= "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"& bankrn &"' and AGNRN='"&agn_rn&"'"
							Query.Open
							agn_strcode 	= Query.FieldByname("STRCODE").value
						end if
					end if
				end if
				
				'������ ��� ��� �������� �� ��� ���� 1�
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='�"&nodeNode.selectSingleNode("�����������").text&"' and docs_prop_rn='104582941' and unitcode='TypeOpersPay'"        ' 104582941 - ��� �������� "��� SAP"
				Query.Open
				Query.SQL.Text="select TYPOPER_MNEMO from DICTOPER where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				typeoper_mnemo = Query.FieldByname("TYPOPER_MNEMO").value
				Query.Close
				
				'������� ������������ ������ �� �� ����
				Query.SQL.Text = "select INTCODE from curnames where curcode='"&nodeNode.selectSingleNode("���������������").text&"'"
				Query.Open
				If Query.IsEmpty then
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ������ � ����� "&nodeNode.selectSingleNode("���������������").text&" �� �������� <<"&comment&">> �� �������. ������������ �������� ��-��������� - RUR")
					scurrency = "RUR"
				else
					scurrency = nodeNode.selectSingleNode("���������������").text
				end if
				Query.Close
				
				'������� ������ � ����� ��������� �������� ��
				StoredProc.StoredProcName="P_BANKDOCSACC_INSERT"
				StoredProc.ParamByName("nCOMPANY").value		= 42903			'��� �������������
				StoredProc.ParamByName("nCRN").value			= 104583621		'��� �������� 1�������� ����� �������� � ���������
				StoredProc.ParamByName("SBANK_TYPEDOC").value	= "������/�"
				StoredProc.ParamByName("SBANK_PREFDOC").value	= DATE_year
				StoredProc.ParamByName("SBANK_NUMBDOC").value	= banknumb
				StoredProc.ParamByName("DBANK_DATEDOC").value	= ConvDate(DATEONLY)
				StoredProc.ParamByName("SVALID_TYPEDOC").value	= doctype
				StoredProc.ParamByName("SVALID_NUMBDOC").value	= docpref & delimiter & docnumb
				StoredProc.ParamByName("DVALID_DATEDOC").value	= docdate
				StoredProc.ParamByName("SFROM_NUMB").value		= nodeNode.selectSingleNode("�������").text
				StoredProc.ParamByName("DFROM_DATE").value		= ConvDate(nodeNode.selectSingleNode("������").text)
				StoredProc.ParamByName("SAGENT_FROM").value		= SAGENT_FROM
				StoredProc.ParamByName("SAGENTF_ACC").value		= jur_strcode
				StoredProc.ParamByName("SAGENT_TO").value		= agn_abbr
				StoredProc.ParamByName("SAGENTT_ACC").value		= agn_strcode
				StoredProc.ParamByName("STYPE_OPER").value		= typeoper_mnemo
				StoredProc.ParamByName("SPAY_INFO").value		= nodeNode.selectSingleNode("�����������������").text
				StoredProc.ParamByName("SPAY_NOTE").value		= SPAY_NOTE
				StoredProc.ParamByName("NPAY_SUM").value		= Replace(nodeNode.selectSingleNode("��������������").text, ".", ",")
				StoredProc.ParamByName("NTAX_SUM").value		= Tax
				StoredProc.ParamByName("NPERCENT_TAX_SUM").value= TAXrate
				StoredProc.ParamByName("SCURRENCY").value		= scurrency
				StoredProc.ParamByName("SJUR_PERS").value		= "�� ���"
				StoredProc.ParamByName("NUNALLOTTED_SUM").value	= 0
				StoredProc.ParamByName("NIS_ADVANCE").value		= 0
				StoredProc.ExecProc
				newRN = StoredProc.ParamByName("nRN").value
				
				'������� ������ � �������� ���������
				StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         '������1�
				StoredProc.ParamByName("PROPERTY").value="������1�"
				StoredProc.ParamByName("UNITCODE").value="BankDocuments"
				StoredProc.ParamByName("RN_SOTR").value=newRN
				StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("������").text)
				StoredProc.ParamByName("NUM_VAL").value=NULL
				StoredProc.ExecProc
			else
				MyFile.Write("INFO "&now()&vbTab&" � ����� ������ �������� �������� �� � ����� "&trim(nodeNode.selectSingleNode("������").text)&"  "&nodeNode.selectSingleNode("�������").text&" "&nodeNode.selectSingleNode("������").text&" - ��������� ��������."&vbNewLine)
			End if
			Query.Close
		next
		MyFile.Write("INFO "&now()&vbTab&" ��������� �������� � ��������� ������ ���������. ����� ����������: "& object_counter & vbNewLine&vbNewLine)
	end if
	
	'������� ���� "��������������������������"'
	Set Incoming = xmlParser.selectNodes("//��������������������������/������")
	If Incoming.length > 0 then
		object_counter = 0
		For Each nodeNode In Incoming
			DATE_array		= NULL
			DATE_year		= NULL
			banknumb		= NULL
			DATETIME_array	= NULL
			DATEONLY		= NULL
			docpref 		= NULL
			docnumb 		= NULL
			doctype 		= NULL
			docdate 		= NULL
			jur_strcode		= NULL
			agn_abbr		= NULL
			agnaccbik		= NULL
			agnacc			= NULL
			agn_strcode		= NULL
			typeoper_mnemo	= NULL
			scurrency		= NULL
			newRN			= NULL
			agn_rn			= NULL
			SPAY_NOTE		= NULL
			bankrn			= NULL
			
			object_counter = object_counter + 1
			
			Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value like '%"&trim(nodeNode.selectSingleNode("������").text)&"%' and docs_prop_rn='104582883' and unitcode='BankDocuments'"
			Query.Open
			If Query.IsEmpty then	
				MyFile.Write("INFO "&now()&vbTab&" � ����� �� ������ �������� ����������� �� � ����� "&trim(nodeNode.selectSingleNode("������").text)&" "&nodeNode.selectSingleNode("�������").text&" "&nodeNode.selectSingleNode("������").text&" - ������ ����� ��������."&vbNewLine)
				
				'������� ��� �� ����
				DATE_array	= Split(nodeNode.selectSingleNode("������").text, "-")
				DATE_year	= DATE_array(0)
				
				'������� ���������� �����
				StoredProc.StoredProcName="P_BANKDOCS_GETNEXTNUMB"
				StoredProc.ParamByName("NCOMPANY").value		= 42903
				StoredProc.ParamByName("SJUR_PERS").value		= "�� ���"
				StoredProc.ParamByName("DBANK_DOCDATE").value	= ConvDate(nodeNode.selectSingleNode("������").text)
				StoredProc.ParamByName("SBANK_DOCTYPE").value	= "�����/�"
				StoredProc.ParamByName("SBANK_DOCPREF").value	= DATE_year
				StoredProc.ExecProc
				banknumb = StoredProc.ParamByName("SBANK_NUMB").value
				
				'������� ���� �� ������ ���� ����������
				DATETIME_array	= Split(nodeNode.selectSingleNode("����").text, "T")
				DATEONLY = DATETIME_array(0)
				
				'������� ������ ��������, ���� � ��������� ������ 1 �������
				docpref = NULL
				docnumb = NULL
				doctype = NULL
				docdate = NULL
				delimiter = NULL
				TAXrate	= 0
				Tax		= 0
				Set TableRows = nodeNode.selectNodes("��������/������")
				If TableRows.length = 1 then
					Set Node = TableRows.nextNode()
					Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&Node.selectSingleNode("�������").text&"' and docs_prop_rn='87456099' and unitcode='Contracts'"        ' 87456099 - ��� �������� "������1�"
					Query.Open
					If not Query.IsEmpty then
						Query.SQL.Text="select DOC_TYPE, DOC_PREF, DOC_NUMB, DOC_DATE from CONTRACTS where RN='"&Query.FieldByname("UNIT_RN").value&"'"
						Query.Open
						docdate = Query.FieldByname("DOC_DATE").value
						docpref = LTrim(Query.FieldByname("DOC_PREF").value)
						docnumb = LTrim(Query.FieldByname("DOC_NUMB").value)
						delimiter = "-"
						Query.SQL.Text="select DOCCODE from DOCTYPES where RN='"&Query.FieldByname("DOC_TYPE").value&"'"
						Query.Open
						doctype = Query.FieldByname("DOCCODE").value
						Query.Close
					end if
					
					'������� ������ ��� � �� ��������
					If Node.selectSingleNode("���������").text="������" or Node.selectSingleNode("���������").text="" then
						TAXrate	= 0
						Tax		= 0
					else
						TAXrate	= Mid(Node.selectSingleNode("���������").text, 4)
						Tax		= Replace(Node.selectSingleNode("��������").text, ".", ",")
					End if
				end if							
				
				'������ �������� ����� ����������� �� �� ���� SAP
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - ��� �������� "��� SAP" (1�������� ����� �������� � ���������)
				Query.Open
				Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				SAGENT_TO = Query.FieldByname("AGNABBR").value
				Query.Close
				
				'������� ����� ������ ����������� ����� ����������� ����� ����� ����� �����
				Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & Trim(nodeNode.selectSingleNode("������������������").text) & "'"
				Query.Open
				If not Query.IsEmpty then
					Query.SQL.Text	= "select STRCODE from AGNACC where AGNACC='"& Trim(nodeNode.selectSingleNode("���������������").text) &"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='5805775'"
					Query.Open
					jur_strcode		= Query.FieldByname("STRCODE").value
					Query.Close
				else
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ��� ����������� ����� � ������� "&nodeNode.selectSingleNode("���������������").text&" � ����� �� ������ ��� "&nodeNode.selectSingleNode("������������������").text&" - ����� ������� ��� �������."&vbNewLine)
				end if
				
				'���� ����������� � ���� � ����������� �� ��������� �������
				if nodeNode.selectSingleNode("�����������").text = "��������������������" then
					'���������� - ���
					Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
					Query.Open
					agn_rn = Query.FieldByname("UNIT_RN").value
					Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value
					
					if nodeNode.selectSingleNode("����������").text="" or nodeNode.selectSingleNode("���������������").text = "" then
						SPAY_NOTE = "���������� (��� SAP) "&nodeNode.selectSingleNode("����������").text&", ���� "&nodeNode.selectSingleNode("���������������").text
						agn_abbr = NULL
						agn_strcode = NULL
					else
						'C��� - �� ��� ������
						agnaccbik		= Trim(nodeNode.selectSingleNode("������������������").text)
						agnacc			= Trim(nodeNode.selectSingleNode("���������������").text)
						If not nodeNode.selectSingleNode("������������������").text="" then
							Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & agnaccbik & "'"
							Query.Open
							bankrn = Query.FieldByname("NRN").value
						else
							bankrn = NULL
						end if
						
						If not IsNull(bankrn) then
							Query.SQL.Text 	= "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"&bankrn&"' and AGNRN='"&agn_rn&"'"
							Query.Open
							If not Query.IsEmpty then
								agn_strcode = Query.FieldByname("STRCODE").value
							else
								'�� ����� �� ������ - �� �������� �����							
								Query.SQL.Text = "select STRCODE from agnacc where agnrn='5805775' and agnbanks='"&bankrn&"'"
								Query.Open
								agn_strcode = Query.FieldByname("STRCODE").value
							end if
						else
							MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ��� ����������� ����� � ������� "&nodeNode.selectSingleNode("���������������").text&" � ����� �� ������ ��� "&nodeNode.selectSingleNode("������������������").text&" - ����� ������� ��� �������."&vbNewLine)
						end if
					end if		
				else
					if nodeNode.selectSingleNode("����������").text = "" then
						SPAY_NOTE = "���������� (��� SAP) "&nodeNode.selectSingleNode("����������").text&", ���� "&nodeNode.selectSingleNode("���������������").text
						agn_abbr	= NULL
						agn_strcode = NULL
					else
						'���� �����������
						Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='"&nodeNode.selectSingleNode("����������").text&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
						Query.Open
						agn_rn = Query.FieldByname("UNIT_RN").value
						Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
						Query.Open
						agn_abbr = Query.FieldByname("AGNABBR").value
						
						'���� ����
						If nodeNode.selectSingleNode("���������������").text = "" then	'������ ���������� ����
							Query.SQL.Text 	= "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
							Query.Open
							agn_strcode = Query.FieldByname("STRCODE").value
						else
							agnaccbik		= Trim(nodeNode.selectSingleNode("������������������").text)
							agnacc			= Trim(nodeNode.selectSingleNode("���������������").text)
							If not nodeNode.selectSingleNode("������������������").text="" then
								Query.SQL.Text	= "select NRN from v_agnbanks where SCODE='����_" & agnaccbik & "'"
								Query.Open
								bankrn = Query.FieldByname("NRN").value
							else
								bankrn = NULL
							end if
							If IsNull(bankrn) then
								Query.SQL.Text 	= "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS is NULL and AGNRN='"&agn_rn&"'"
								Query.Open
								agn_strcode 	= Query.FieldByname("STRCODE").value
							else
								Query.SQL.Text 	= "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"& bankrn &"' and AGNRN='"&agn_rn&"'"
								Query.Open
								agn_strcode 	= Query.FieldByname("STRCODE").value
							end if
						end if
					end if
				end if
								
				'������ ��� ��� �������� �� ��� ���� 1�
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='�"&nodeNode.selectSingleNode("�����������").text&"' and docs_prop_rn='104582941' and unitcode='TypeOpersPay'"        ' 104582941 - ��� �������� "��� SAP"
				Query.Open
				Query.SQL.Text="select TYPOPER_MNEMO from DICTOPER where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				typeoper_mnemo = Query.FieldByname("TYPOPER_MNEMO").value
				Query.Close
				
				'������� ������������ ������ �� �� ����
				Query.SQL.Text = "select INTCODE from curnames where curcode='"&nodeNode.selectSingleNode("���������������").text&"'"
				Query.Open
				If Query.IsEmpty then
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" ������ � ����� "&nodeNode.selectSingleNode("���������������").text&" �� �������� <<"&comment&">> �� �������. ������������ �������� ��-��������� - RUR")
					scurrency = "RUR"
				else
					scurrency = nodeNode.selectSingleNode("���������������").text
				end if
				Query.Close
								
				'������� ������ � ����� ��������� ����������� ��
				StoredProc.StoredProcName="P_BANKDOCSACC_INSERT"
				StoredProc.ParamByName("nCOMPANY").value		= 42903			'��� �������������
				StoredProc.ParamByName("nCRN").value			= 104583621		'��� �������� 1�������� ����� �������� � ���������
				StoredProc.ParamByName("SBANK_TYPEDOC").value	= "�����/�"
				StoredProc.ParamByName("SBANK_PREFDOC").value	= DATE_year
				StoredProc.ParamByName("SBANK_NUMBDOC").value	= banknumb
				StoredProc.ParamByName("DBANK_DATEDOC").value	= ConvDate(DATEONLY)
				StoredProc.ParamByName("SVALID_TYPEDOC").value	= doctype
				StoredProc.ParamByName("SVALID_NUMBDOC").value	= docpref & delimiter & docnumb
				StoredProc.ParamByName("DVALID_DATEDOC").value	= docdate
				StoredProc.ParamByName("SFROM_NUMB").value		= nodeNode.selectSingleNode("�������").text
				StoredProc.ParamByName("DFROM_DATE").value		= ConvDate(nodeNode.selectSingleNode("������").text)
				StoredProc.ParamByName("SAGENT_FROM").value		= agn_abbr
				StoredProc.ParamByName("SAGENTF_ACC").value		= agn_strcode
				StoredProc.ParamByName("SAGENT_TO").value		= SAGENT_TO
				StoredProc.ParamByName("SAGENTT_ACC").value		= jur_strcode
				StoredProc.ParamByName("STYPE_OPER").value		= typeoper_mnemo
				StoredProc.ParamByName("SPAY_INFO").value		= nodeNode.selectSingleNode("�����������������").text 
				StoredProc.ParamByName("SPAY_NOTE").value		= SPAY_NOTE
				StoredProc.ParamByName("NPAY_SUM").value		= Replace(nodeNode.selectSingleNode("��������������").text, ".", ",")
				StoredProc.ParamByName("NTAX_SUM").value		= Tax
				StoredProc.ParamByName("NPERCENT_TAX_SUM").value= TAXrate
				StoredProc.ParamByName("SCURRENCY").value		= scurrency
				StoredProc.ParamByName("SJUR_PERS").value		= "�� ���"
				StoredProc.ParamByName("NUNALLOTTED_SUM").value	= 0
				StoredProc.ParamByName("NIS_ADVANCE").value		= 0
				StoredProc.ExecProc
				newRN = StoredProc.ParamByName("nRN").value
				
				'������� ������ � �������� ���������
				StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         '������1�
				StoredProc.ParamByName("PROPERTY").value="������1�"
				StoredProc.ParamByName("UNITCODE").value="BankDocuments"
				StoredProc.ParamByName("RN_SOTR").value=newRN
				StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("������").text)
				StoredProc.ParamByName("NUM_VAL").value=NULL
				StoredProc.ExecProc
			else
				MyFile.Write("INFO "&now()&vbTab&" � ����� ������ �������� �������� �� � ����� "&trim(nodeNode.selectSingleNode("������").text)&"  "&nodeNode.selectSingleNode("�������").text&" "&nodeNode.selectSingleNode("������").text&" - ��������� ��������."&vbNewLine)
			End if
			Query.Close
		next
		MyFile.Write("INFO "&now()&vbTab&" ��������� ����������� �� ��������� ����� ���������.  ����� ����������: "& object_counter & vbNewLine)
	end if
	
	'������� ���� ������
	Set oldReply = CreateObject("Msxml2.DOMDocument")
	oldReply.async = False
	oldReply.load "\\10.130.32.52\Tatneft\Mess_20100_UH.xml"
	' ��������� �� ������ ��������
	If oldReply.parseError.errorCode Then
			MsgBox oldReply.parseError.Reason
	End If
	If not oldReply.parseError.errorCode = -2146697210 then
		oldReplyNumber = cInt(oldReply.selectSingleNode("/������/���������������������������").text)+1
	else
		oldReplyNumber = 1
	End if
	Set newReply = CreateObject("Msxml2.DOMDocument")
	newReply.appendChild(newReply.createProcessingInstruction("xml", "version='1.0' encoding='utf-8'"))
	Set rootNode = newReply.appendChild( newReply.createElement("������") )
	rootNode.setAttribute "xmlns", "http://localhost/ExchangeUH_FileResponse"
	rootNode.setAttribute "xmlns:xs", "http://www.w3.org/2001/XMLSchema"
	rootNode.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
	Set subNode = rootNode.appendChild(newReply.createElement("���������������������������"))
	subNode.text = oldReplyNumber
	Set subNode = rootNode.appendChild(newReply.createElement("�����������������������"))
	subNode.text = xmlParser.selectSingleNode("//���������������������������").text
	newReply.save("\\10.130.32.52\Tatneft\Mess_20100_UH.xml")
	newReply.save("\\10.130.32.52\Tatneft\Exch_logs\Mess_20100_UH.xml")
	
	'������ ��������
	MyFile.Write(vbTab&now()&vbTab&" ������ ������ �� 1�:�� � ��� ����� ������� ��������."&vbNewLine)
	MsgBox "������ ������ ��������"&vbNewLine&"����������� � ������� 1C-Parus_exchange.log"
end sub

Function ShowError(XMLDOMParseError)
	mess = _
	"parseError.errorCode: " & XMLDOMParseError.errorCode & vbCrLf & _
	"parseError.filepos: " & XMLDOMParseError.filepos & vbCrLf & _
	"parseError.line: " & XMLDOMParseError.line & vbCrLf & _
	"parseError.linepos: " & XMLDOMParseError.linepos & vbCrLf & _
	"parseError.reason: " & XMLDOMParseError.reason & vbCrLf & _
	"parseError.srcText: " & XMLDOMParseError.srcText & vbCrLf & _
	"parseError.url: " & XMLDOMParseError.url & vbCrLf
	Msgbox mess
End Function

Function ConvDate(DateToConvert)
	DATE_array	= Split(DateToConvert, "-")
	DATE_year	= DATE_array(0)
	DATE_month	= DATE_array(1)
	DATE_day	= DATE_array(2)
	ConvDate	= DATE_day&"."&DATE_month&"."&DATE_year
End Function

Function GetFaceAccCat(subdiv)
	Select Case subdiv
		Case "�����.01"
			CatRN = 4471103 
		Case "�����.02"
		Case "�����.03"
		Case "�����.04"
		Case else
	End Select
	Query.SQL.Text="select NAME from ACATALOG where RN='"&CatRN&"'"
	Query.Open
	GetFaceAccCat = "test" 'Query.FieldByname("NAME").value
End Function