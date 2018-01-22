sub xml_import
	Dim contragents 'коллекцию XMLDOMNodeList всех элементов заданного типа
	Dim cAgent 'элемент коллекции Контрагенты
	Dim agncounter

	' Загружаем XML-документ
	Set xmlParser = CreateObject("Msxml2.DOMDocument")
	xmlParser.async = False
	xmlParser.load "\\10.130.32.52\Tatneft\Mess_UH_20100.xml"

	' Проверяем на ошибки загрузки
	If xmlParser.parseError.errorCode Then
					MsgBox xmlParser.parseError.Reason
	End If

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.OpenTextFile("\\10.130.32.52\Tatneft\Exch_logs\1C-Parus_exchange.log", 8, True)
	MyFile.Write(vbNewLine&"**********************************************************************************************************************************"&vbNewLine&vbTab&now()&vbTab&" Начало импорта данных из 1С:УХ в ИСУ Парус"&vbNewLine)

	' Находим узел Контрагенты
	Set contragents = xmlParser.selectNodes("//Контрагенты/Строки")
	If contragents.length > 0 then
		' ПЕРЕБИРАЕМ СПИСОК КОНТРАГЕНТОВ В XML-ДОКУМЕНТЕ
		object_counter=0
		T_Agn_error = False
		For Each nodeNode In contragents
			SAPcode                                = NULL
			INN                                        = NULL
			KPP                                        = NULL
			agntype                                = NULL
			agncounter                        = NULL
			DoubledAgentsString        = NULL
			SAPcode                                = NULL
			newRN                                = NULL
			AGNABBR                                = NULL
			divider                                = NULL

			object_counter=object_counter+1

			SAPcode = RTrim(nodeNode.selectSingleNode("Код").text)
			INN = nodeNode.selectSingleNode("ИНН").text
			KPP = nodeNode.selectSingleNode("КПП").text

			if INN="1651057954" then
				divider="/"
			else
				divider="^"
			end if

			If nodeNode.selectSingleNode("ЮридическоеФизическоеЛицо").text = "ЮридическоеЛицо" Then
				agntype=0
			Else
				agntype=1
			end If

			' НАЙДЕМ RN КОНТРАГЕНТА ПО КОДУ SAP В ТАБЛИЦЕ СВОЙСТВ ОБЪЕКТОВ
			Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='"&SAPcode&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - код свойства "Код SAP" (1ИЗМЕНИТЬ ПОСЛЕ ПЕРЕНОСА В ПРОДУКТИВ)
			Query.Open
			If Query.IsEmpty then
				' ВДРУГ У НАС УЖЕ ЕСТЬ КОНТРАГЕНТ С ТАКИМ ЖЕ ИНН/КПП НО БЕЗ КОДА SAP
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
					T_New_agn = True
					MyFile.Write("INFO "&now()&vbTab&" В Парус не найден контрагент с кодом SAP "&SAPcode&" или с парой ИНН/КПП "&INN&"/"&KPP&" ("&nodeNode.selectSingleNode("Наименование").text&") - создаю нового контрагента."&vbNewLine)
				else
					agn_name = CaQuery.FieldByname("AGNNAME").value
					' ИНН/КПП СОВПАДАЕТ, А КОД SAP НЕ ЗАПОЛНЕН - СТАВИМ ТРИГГЕР ДЛЯ ОСТАНОВА В КОНЦЕ БЛОКА ЗАГРУЗКИ КОНТРАГЕНТОВ
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" В таблице найден контрагент с такими же ИНН/КПП, но без кода SAP "&SAPcode&" (либо с другим кодом SAP) - контрагент не создан, необходимо проверить вручную: "&INN&"/"&KPP&" ("&agn_name&")"&vbNewLine)
					T_Agn_error = True
					T_New_agn = False
				end if

				If T_New_agn then
					AGNABBR = INN&divider&KPP
					REM Query.SQL.Text="select AGNABBR, STR_VALUE from AGNLIST a, DOCS_PROPS_VALS b where a.AGNABBR='"&AGNABBR&"' and a.RN=b.UNIT_RN and DOCS_PROP_RN=105510718 and not STR_VALUE='"&SAPcode&"'"
					REM Query.Open
					REM If not Query.IsEmpty then        'это дубль, во избежание ошибок нарушения уникальности сохраним в качестве мнемокода код SAP
					REM MyFile.Write("INFO "&now()&vbTab&" В Парус найден контрагент с мнемокодом "&AGNABBR&", но другим кодом SAP "&Query.FieldByname("STR_VALUE")&" ("&nodeNode.selectSingleNode("Наименование").text&") - в качестве мнемокода будет указан код SAP "&SAPcode&"."&vbNewLine)
					REM AGNABBR=SAPcode
					REM end if

					'ДОБАВИМ ЗАПИСЬ О НОВОМ КОНТРАГЕНТЕ
					StoredProc.StoredProcName="P_AGNLIST_INSERT"
					StoredProc.ParamByName("nCOMPANY").value=42903                                        'код подразделения
					StoredProc.ParamByName("CRN").value=155332                                                'код каталога
					StoredProc.ParamByName("AGNTYPE").value=agntype
					StoredProc.ParamByName("AGNNAME").value=nodeNode.selectSingleNode("Наименование").text
					StoredProc.ParamByName("sFULLNAME").value=nodeNode.selectSingleNode("НаименованиеПолное").text
					StoredProc.ParamByName("AGNIDNUMB").value=INN
					StoredProc.ParamByName("sREASON_CODE").value=nodeNode.selectSingleNode("КПП").text
					StoredProc.ParamByName("sOGRN").value=nodeNode.selectSingleNode("ОГРН").text
					StoredProc.ParamByName("ORGCODE").value=nodeNode.selectSingleNode("ОКПО").text
					StoredProc.ParamByName("AGNABBR").value=AGNABBR
					StoredProc.ParamByName("PHONE").value=nodeNode.selectSingleNode("Телефон").text
					StoredProc.ParamByName("EMP").value=0
					StoredProc.ParamByName("nSEX").value=0
					StoredProc.ParamByName("nRESIDENT_SIGN").value=0
					StoredProc.ParamByName("nCOEFFIC").value=0
					StoredProc.ParamByName("nIND_BUSINESSMAN").value=0
					StoredProc.ExecProc
					newRN = StoredProc.ParamByName("nRN").value

					' ЗАПИШЕМ КОД SAP В СВОЙСТВО ОБЪЕКТА
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" 'процедурка для записи доп свойства
					StoredProc.ParamByName("PROPERTY").value="КодКонтр1С"
					StoredProc.ParamByName("UNITCODE").value="AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value=newRN
					StoredProc.ParamByName("ST_VAL").value=SAPcode
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc

					'ДОБАВИМ ЗАПИСЬ ОБ АДРЕСЕ НОВОГО КОНТРАГЕНТА
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" 'процедурка для записи доп свойства
					StoredProc.ParamByName("PROPERTY").value="АдрСтрокЮрид1С"
					StoredProc.ParamByName("UNITCODE").value="AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value=newRN
					StoredProc.ParamByName("ST_VAL").value=nodeNode.selectSingleNode("АдресЮридический").text
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" 'процедурка для записи доп свойства
					StoredProc.ParamByName("PROPERTY").value="АдрСтрокПочт1С"
					StoredProc.ParamByName("UNITCODE").value="AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value=newRN
					StoredProc.ParamByName("ST_VAL").value=nodeNode.selectSingleNode("АдресПочтовый").text
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc
				end If
			else        'найдена запись с нашим кодом SAP
				' НАЙДЕМ КОНТРАГЕНТА В ТАБЛИЦЕ ПО RN
				CaQuery=Query
				CaQuery.Sql.Text = "select * from AGNLIST where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				CaQuery.Open

				' ПОДСЧИТАЕМ КОЛИЧЕСТВО СОВПАВШИХ КОНТРАГЕНТОВ И СОСТАВИМ ИХ СПИСОК НА ВСЯКИЙ СЛУЧАЙ
				agncounter = 0
				DoubledAgentsString=""
				do while not CaQuery.EOF
					agncounter=agncounter+1
					DoubledAgentsString=DoubledAgentsString & vbTab&agncounter & ") " & CaQuery.FieldByname("AGNIDNUMB").value & "/" & CaQuery.FieldByname("REASON_CODE").value & " " & CaQuery.FieldByname("AGNNAME").value & vbNewLine
					CaQuery.next
				loop
				If agncounter>1 then
					'РУГНЕМСЯ НА ДУБЛИРОВАНИЕ КОДОВ SAP И СДЕЛАЕМ ОТМЕТКУ В ЛОГЕ
					MyFile.Write("ERROR "&now()&vbTab&" В Парус найдено несколько контрагентов с кодом SAP "&SAPcode&" ("&nodeNode.selectSingleNode("Наименование").text&"),  - необходимо проверить вручную, объект не обновлен:"&vbNewLine)
					T_Agn_error = True
					MyFile.Write(DoubledAgentsString)
				elseIf agncounter=1 Then
					MyFile.Write("INFO "&now()&vbTab&" В Парус найден контрагент с кодом SAP "&SAPcode&" ("&nodeNode.selectSingleNode("Наименование").text&") - обновляю данные контрагента."&vbNewLine)
					RN					= CaQuery.FieldByname("RN").value
					ECONCODE			= CaQuery.FieldByname("ECONCODE").value
					If nodeNode.selectSingleNode("ОКПО").text = "" then
						ORGCODE			= CaQuery.FieldByname("ORGCODE").value
					else
						ORGCODE			= nodeNode.selectSingleNode("ОКПО").text
					end if
					AGNFAMILYNAME		= CaQuery.FieldByname("AGNFAMILYNAME").value
					AGNFIRSTNAME		= CaQuery.FieldByname("AGNFIRSTNAME").value
					AGNLASTNAME			= CaQuery.FieldByname("AGNLASTNAME").value
					AGNFAMILYNAME_TO	= CaQuery.FieldByname("AGNFAMILYNAME_TO").value
					AGNFIRSTNAME_TO    	= CaQuery.FieldByname("AGNFIRSTNAME_TO").value
					AGNLASTNAME_TO     	= CaQuery.FieldByname("AGNLASTNAME_TO").value
					AGNFAMILYNAME_FR  	= CaQuery.FieldByname("AGNFAMILYNAME_FR").value
					AGNFIRSTNAME_FR    	= CaQuery.FieldByname("AGNFIRSTNAME_FR").value
					AGNLASTNAME_FR    	= CaQuery.FieldByname("AGNLASTNAME_FR").value
					AGNFAMILYNAME_AC  	= CaQuery.FieldByname("AGNFAMILYNAME_AC").value
					AGNFIRSTNAME_AC   	= CaQuery.FieldByname("AGNFIRSTNAME_AC").value
					AGNLASTNAME_AC    	= CaQuery.FieldByname("AGNLASTNAME_AC").value
					AGNFAMILYNAME_ABL 	= CaQuery.FieldByname("AGNFAMILYNAME_ABL").value
					AGNFIRSTNAME_ABL  	= CaQuery.FieldByname("AGNFIRSTNAME_ABL").value
					AGNLASTNAME_ABL   	= CaQuery.FieldByname("AGNLASTNAME_ABL").value
					EMPPOST           	= CaQuery.FieldByname("EMPPOST").value
					EMPPOST_FROM       	= CaQuery.FieldByname("EMPPOST_FROM").value
					EMPPOST_TO         	= CaQuery.FieldByname("EMPPOST_TO").value
					EMPPOST_AC        	= CaQuery.FieldByname("EMPPOST_AC").value
					EMPPOST_ABL      	= CaQuery.FieldByname("EMPPOST_ABL").value
					AGNBURN            	= CaQuery.FieldByname("AGNBURN").value
					If nodeNode.selectSingleNode("Телефон").text = "" then
						PHONE           = CaQuery.FieldByname("PHONE").value
					else
						PHONE         	= nodeNode.selectSingleNode("Телефон").text
					end if
					PHONE2          	= CaQuery.FieldByname("PHONE2").value
					FAX               	= CaQuery.FieldByname("FAX").value
					TELEX            	= CaQuery.FieldByname("TELEX").value
					If nodeNode.selectSingleNode("ЭлектроннаяПочта").text = "" then
						MAIL        	= CaQuery.FieldByname("MAIL").value
					else
						MAIL         	= nodeNode.selectSingleNode("ЭлектроннаяПочта").text
					end if
					IMAGE          		= CaQuery.FieldByname("IMAGE").value
					DISCDATE        	= CaQuery.FieldByname("DISCDATE").value
					AGN_COMMENT     	= CaQuery.FieldByname("AGN_COMMENT").value
					PENSION_NBR        	= CaQuery.FieldByname("PENSION_NBR").value
					MEDPOLICY_SER      	= CaQuery.FieldByname("MEDPOLICY_SER").value
					MEDPOLICY_NUMB    	= CaQuery.FieldByname("MEDPOLICY_NUMB").value
					PROPFORM          	= CaQuery.FieldByname("PROPFORM").value
					TAXPSTATUS        	= CaQuery.FieldByname("TAXPSTATUS").value
					PRFMLSTS          	= CaQuery.FieldByname("PRFMLSTS").value
					PRNATION          	= CaQuery.FieldByname("PRNATION").value
					CITIZENSHIP       	= CaQuery.FieldByname("CITIZENSHIP").value
					'CITIZENOKIN       	= CaQuery.FieldByname("CITIZENOKIN").value
					ADDR_BURN          	= CaQuery.FieldByname("ADDR_BURN").value
					PRMLREL           	= CaQuery.FieldByname("PRMLREL").value
					OKATO           	= CaQuery.FieldByname("OKATO").value
					PFR_NAME      		= CaQuery.FieldByname("PFR_NAME").value
					PFR_FILL_DATE   	= CaQuery.FieldByname("PFR_FILL_DATE").value
					PFR_REG_DATE     	= CaQuery.FieldByname("PFR_REG_DATE").value
					PFR_REG_NUMB      	= CaQuery.FieldByname("PFR_REG_NUMB").value
					If nodeNode.selectSingleNode("ОГРН").text = "" then
						OGRN         	= CaQuery.FieldByname("OGRN").value
					else
						OGRN         	= nodeNode.selectSingleNode("ОГРН").text
					end if
					OKFS              	= CaQuery.FieldByname("OKFS").value
					If nodeNode.selectSingleNode("ОКОПФ").text = "" then
						OKOPF         	= CaQuery.FieldByname("OKOPF").value
					else
						OKOPF         	= nodeNode.selectSingleNode("ОКОПФ").text
					end if
					TFOMS           	= CaQuery.FieldByname("TFOMS").value
					FSS_REG_NUMB      	= CaQuery.FieldByname("FSS_REG_NUMB").value
					FSS_SUBCODE        	= CaQuery.FieldByname("FSS_SUBCODE").value
					AGNDEATH         	= CaQuery.FieldByname("AGNDEATH").value
					OKTMO            	= CaQuery.FieldByname("OKTMO").value
					INN_CITIZENSHIP   	= CaQuery.FieldByname("INN_CITIZENSHIP").value

					'УСТАНОВИМ В NULL ВСЕ ЗНАЧЕНИЯ 0
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

					AGNABBR = nodeNode.selectSingleNode("ИНН").text&divider&nodeNode.selectSingleNode("КПП").text
					Query.SQL.Text="select AGNABBR, STR_VALUE from AGNLIST a, DOCS_PROPS_VALS b where a.AGNABBR='"&AGNABBR&"' and a.RN=b.UNIT_RN and DOCS_PROP_RN=105510718 and not STR_VALUE='"&SAPcode&"'"
					Query.Open
					If not Query.IsEmpty then        'это дубль, во избежание ошибок нарушения уникальности изменим мнемокод
						MyFile.Write("INFO "&now()&vbTab&" В Парус найден контрагент с мнемокодом "&AGNABBR&", но другим кодом SAP "&Query.FieldByname("STR_VALUE").value&" ("&nodeNode.selectSingleNode("Наименование").text&") - в качестве разделителя мнемокода будет указан символ v "&SAPcode&"."&vbNewLine)
						divider="v"
					end if

					'ОБНОВИМ ЗАПИСЬ О КОНТРАГЕНТЕ
					StoredProc.StoredProcName="P_AGNLIST_UPDATE"
					StoredProc.ParamByName("nCOMPANY").value                = 42903        'код подразделения
					StoredProc.ParamByName("RN").value                                = RN
					StoredProc.ParamByName("AGNABBR").value                        = nodeNode.selectSingleNode("ИНН").text&divider&nodeNode.selectSingleNode("КПП").text
					StoredProc.ParamByName("AGNTYPE").value                        = agntype
					StoredProc.ParamByName("AGNNAME").value                        = nodeNode.selectSingleNode("Наименование").text
					StoredProc.ParamByName("AGNIDNUMB").value                = nodeNode.selectSingleNode("ИНН").text
					StoredProc.ParamByName("ECONCODE").value                = ECONCODE
					StoredProc.ParamByName("ORGCODE").value                        = nodeNode.selectSingleNode("ОКПО").text
					StoredProc.ParamByName("AGNFAMILYNAME").value        = AGNFAMILYNAME
					StoredProc.ParamByName("AGNFIRSTNAME").value        = AGNFIRSTNAME
					StoredProc.ParamByName("AGNLASTNAME").value                = AGNLASTNAME
					StoredProc.ParamByName("AGNFAMILYNAME_TO").value= AGNFAMILYNAME_TO
					StoredProc.ParamByName("AGNFIRSTNAME_TO").value        = AGNFIRSTNAME_TO
					StoredProc.ParamByName("AGNLASTNAME_TO").value        = AGNLASTNAME_TO
					StoredProc.ParamByName("AGNFAMILYNAME_FR").value= AGNFAMILYNAME_FR
					StoredProc.ParamByName("AGNFIRSTNAME_FR").value        = AGNFIRSTNAME_FR
					StoredProc.ParamByName("AGNLASTNAME_FR").value        = AGNLASTNAME_FR
					StoredProc.ParamByName("AGNFAMILYNAME_AC").value= AGNFAMILYNAME_AC
					StoredProc.ParamByName("AGNFIRSTNAME_AC").value        = AGNFIRSTNAME_AC
					StoredProc.ParamByName("AGNLASTNAME_AC").value        = AGNLASTNAME_AC
					StoredProc.ParamByName("AGNFAMILYNAME_ABL").value= AGNFAMILYNAME_ABL
					StoredProc.ParamByName("AGNFIRSTNAME_ABL").value= AGNFIRSTNAME_ABL
					StoredProc.ParamByName("AGNLASTNAME_ABL").value        = AGNLASTNAME_ABL
					StoredProc.ParamByName("EMP").value                                = 0
					StoredProc.ParamByName("EMPPOST").value                        = EMPPOST
					StoredProc.ParamByName("EMPPOST_FROM").value        = EMPPOST_FROM
					StoredProc.ParamByName("EMPPOST_TO").value                = EMPPOST_TO
					StoredProc.ParamByName("EMPPOST_AC").value                = EMPPOST_AC
					StoredProc.ParamByName("EMPPOST_ABL").value                = EMPPOST_ABL
					StoredProc.ParamByName("AGNBURN").value                        = AGNBURN
					StoredProc.ParamByName("PHONE").value                        = PHONE
					StoredProc.ParamByName("PHONE2").value                        = PHONE2
					StoredProc.ParamByName("FAX").value                                = FAX
					StoredProc.ParamByName("TELEX").value                        = TELEX
					StoredProc.ParamByName("MAIL").value                        = MAIL
					StoredProc.ParamByName("IMAGE").value                        = IMAGE
					StoredProc.ParamByName("dDISCDATE").value                = DISCDATE
					StoredProc.ParamByName("AGN_COMMENT").value                = AGN_COMMENT
					StoredProc.ParamByName("nSEX").value                        = 0
					StoredProc.ParamByName("sPENSION_NBR").value        = PENSION_NBR
					StoredProc.ParamByName("sMEDPOLICY_SER").value        = MEDPOLICY_SER
					StoredProc.ParamByName("sMEDPOLICY_NUMB").value        = MEDPOLICY_NUMB
					StoredProc.ParamByName("sPROPFORM").value                = propform
					StoredProc.ParamByName("sREASON_CODE").value        = nodeNode.selectSingleNode("КПП").text
					StoredProc.ParamByName("nRESIDENT_SIGN").value        = 0
					StoredProc.ParamByName("sTAXPSTATUS").value                = taxpstatus
					StoredProc.ParamByName("sOGRN").value                        = OGRN
					StoredProc.ParamByName("sPRFMLSTS").value                = prfmlsts
					StoredProc.ParamByName("sPRNATION").value                = prnation
					StoredProc.ParamByName("sCITIZENSHIP").value        = citizenship
					'StoredProc.ParamByName("CITIZENOKIN").value                = CITIZENOKIN
					StoredProc.ParamByName("ADDR_BURN").value                = ADDR_BURN
					StoredProc.ParamByName("sPRMLREL").value                = prmlrel
					StoredProc.ParamByName("sOKATO").value                        = OKATO
					StoredProc.ParamByName("sPFR_NAME").value                = PFR_NAME
					StoredProc.ParamByName("dPFR_FILL_DATE").value        = PFR_FILL_DATE
					StoredProc.ParamByName("dPFR_REG_DATE").value        = PFR_REG_DATE
					StoredProc.ParamByName("sPFR_REG_NUMB").value        = PFR_REG_NUMB
					StoredProc.ParamByName("sFULLNAME").value                = nodeNode.selectSingleNode("НаименованиеПолное").text
					StoredProc.ParamByName("sOKFS").value                        = OKFS
					StoredProc.ParamByName("sOKOPF").value                        = OKOPF
					StoredProc.ParamByName("sTFOMS").value                        = TFOMS
					StoredProc.ParamByName("sFSS_REG_NUMB").value        = FSS_REG_NUMB
					StoredProc.ParamByName("sFSS_SUBCODE").value        = FSS_SUBCODE
					StoredProc.ParamByName("nCOEFFIC").value                = 0
					StoredProc.ParamByName("dAGNDEATH").value                = AGNDEATH
					StoredProc.ParamByName("sOKTMO").value                        = OKTMO
					StoredProc.ParamByName("sINN_CITIZENSHIP").value= INN_CITIZENSHIP
					StoredProc.ParamByName("nIND_BUSINESSMAN").value=0
					StoredProc.ExecProc

					'ОБНОВИМ ЗАПИСЬ ОБ АДРЕСЕ КОНТРАГЕНТА
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" 'процедурка для записи доп свойства
					StoredProc.ParamByName("PROPERTY").value        = "АдрСтрокЮрид1С"
					StoredProc.ParamByName("UNITCODE").value        = "AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value                = RN
					StoredProc.ParamByName("ST_VAL").value                = nodeNode.selectSingleNode("АдресЮридический").text
					StoredProc.ParamByName("NUM_VAL").value                = NULL
					StoredProc.ExecProc
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS" 'процедурка для записи доп свойства
					StoredProc.ParamByName("PROPERTY").value        = "АдрСтрокПочт1С"
					StoredProc.ParamByName("UNITCODE").value        = "AGNLIST"
					StoredProc.ParamByName("RN_SOTR").value                = RN
					StoredProc.ParamByName("ST_VAL").value                = nodeNode.selectSingleNode("АдресПочтовый").text
					StoredProc.ParamByName("NUM_VAL").value                = NULL
					StoredProc.ExecProc
				end If
				CaQuery.Close
			end if
			Query.Close
		Next
		MyFile.Write("INFO "&now()&vbTab&" Контрагенты загружены. Всего обработано: "& object_counter & vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" В текущей выгрузке на найдено информации о контрагентах."&vbNewLine)
	End if

	If T_Agn_error = True then
		MsgBox "Загрузка контрагентов прошла с ошибкой. Проверьте журнал 1C-Parus_exchange.log"
		Exit Sub
	end if

	'НАХОДИМ УЗЕЛ "БАНКИ"
	Set Banks = xmlParser.selectNodes("//Банки/Строки")
	If Banks.length > 0 then
		'ПЕРЕБИРАЕМ СПИСОК БАНКОВ
		object_counter=0
		For Each nodeNode In Banks
			object_counter=object_counter+1
			agnbank_rn        = NULL

			BankQuery=Query
			BankQuery.Sql.Text = "select RN, AGNRN from AGNBANKS where BANKFCODEACC='"&nodeNode.selectSingleNode("БИК").text&"' and CRN='104583471'" 'ищем банки только в каталоге, отдельно созданном для банков из 1С
			BankQuery.Open
			agnbank_rn=BankQuery.FieldByname("RN").value
			If BankQuery.IsEmpty Then
				MyFile.Write("INFO "&now()&vbTab&" В Парус не найден банк с кодом БИК "&nodeNode.selectSingleNode("БИК").text&" ("&nodeNode.selectSingleNode("Наименование").text&") - создаю новый банк."&vbNewLine)
				'ДОБАВИМ ЗАПИСЬ О НОВОМ КОНТРАГЕНТЕ-БАНКЕ
				StoredProc.StoredProcName="P_AGNLIST_INSERT"
				StoredProc.ParamByName("nCOMPANY").value=42903                                        'код подразделения
				StoredProc.ParamByName("CRN").value=104582949                                        'код каталога с банками
				StoredProc.ParamByName("AGNTYPE").value=0
				StoredProc.ParamByName("AGNNAME").value=nodeNode.selectSingleNode("Наименование").text&", "&nodeNode.selectSingleNode("Город").text
				StoredProc.ParamByName("AGNABBR").value="БАНК_"&nodeNode.selectSingleNode("БИК").text
				StoredProc.ParamByName("EMP").value=0
				StoredProc.ParamByName("nSEX").value=0
				StoredProc.ParamByName("nRESIDENT_SIGN").value=0
				StoredProc.ParamByName("nCOEFFIC").value=0
				StoredProc.ParamByName("nIND_BUSINESSMAN").value=0
				StoredProc.ExecProc

				'ДОБАВИМ ЗАПИСЬ О НОВОМ ЭЛЕМЕНТЕ СЛОВАРЯ "БАНКОВСКИЕ УЧРЕЖДЕНИЯ"
				StoredProc.StoredProcName="P_AGNBANKS_INSERT"
				StoredProc.ParamByName("nCOMPANY").value=42903                                        'код подразделения
				StoredProc.ParamByName("nCRN").value=104583471                                        'код каталога с банками
				StoredProc.ParamByName("sBANKFCODEACC").value=nodeNode.selectSingleNode("БИК").text
				StoredProc.ParamByName("sBANKACC").value=nodeNode.selectSingleNode("КоррСчет").text
				StoredProc.ParamByName("sCODE").value="БАНК_"&nodeNode.selectSingleNode("БИК").text
				StoredProc.ExecProc
			else
				MyFile.Write("INFO "&now()&vbTab&" В Парус найден банк с кодом БИК "&nodeNode.selectSingleNode("БИК").text&" ("&nodeNode.selectSingleNode("Наименование").text&") - обновляю существующий банк."&vbNewLine)
				bAgnQuery = Query
				bAgnQuery.SQL.Text="select * from agnlist where RN='"&BankQuery.FieldByname("AGNRN").value&"'"
				bAgnQuery.Open

				REM 'ОБНОВИМ ЗАПИСЬ О КОНТРАГЕНТЕ-БАНКЕ
				REM StoredProc.StoredProcName="P_AGNLIST_UPDATE"
				REM StoredProc.ParamByName("RN").value=bAgnQuery.FieldByname("RN").value                                                'код контрагента
				REM StoredProc.ParamByName("AGNABBR").value="БАНК_"&nodeNode.selectSingleNode("БИК").text
				REM StoredProc.ParamByName("AGNTYPE").value=0
				REM StoredProc.ParamByName("AGNNAME").value=nodeNode.selectSingleNode("Наименование").text&", "&nodeNode.selectSingleNode("Город").text
				REM StoredProc.ParamByName("EMP").value=0
				REM StoredProc.ParamByName("nSEX").value=0
				REM StoredProc.ParamByName("nRESIDENT_SIGN").value=0
				REM StoredProc.ParamByName("nCOEFFIC").value=0
				REM StoredProc.ExecProc
				REM bAgnQuery.Close

				'УСТАНОВИМ В NULL ВСЕ ЗНАЧЕНИЯ 0
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
				StoredProc.ParamByName("nCOMPANY").value                = 42903        'код подразделения
				StoredProc.ParamByName("RN").value                                = bAgnQuery.FieldByname("RN").value        'код контрагента
				StoredProc.ParamByName("AGNABBR").value                        = "БАНК_"&nodeNode.selectSingleNode("БИК").text
				StoredProc.ParamByName("AGNTYPE").value                        = 0
				StoredProc.ParamByName("AGNNAME").value                        = nodeNode.selectSingleNode("Наименование").text&", "&nodeNode.selectSingleNode("Город").text
				StoredProc.ParamByName("AGNIDNUMB").value                = bAgnQuery.FieldByname("AGNIDNUMB").value
				StoredProc.ParamByName("ECONCODE").value                = bAgnQuery.FieldByname("ECONCODE").value
				StoredProc.ParamByName("ORGCODE").value                        = bAgnQuery.FieldByname("ORGCODE").value
				StoredProc.ParamByName("AGNFAMILYNAME").value        = bAgnQuery.FieldByname("AGNFAMILYNAME").value
				StoredProc.ParamByName("AGNFIRSTNAME").value        = bAgnQuery.FieldByname("AGNFIRSTNAME").value
				StoredProc.ParamByName("AGNLASTNAME").value                = bAgnQuery.FieldByname("AGNLASTNAME").value
				StoredProc.ParamByName("AGNFAMILYNAME_TO").value= bAgnQuery.FieldByname("AGNFAMILYNAME_TO").value
				StoredProc.ParamByName("AGNFIRSTNAME_TO").value        = bAgnQuery.FieldByname("AGNFIRSTNAME_TO").value
				StoredProc.ParamByName("AGNLASTNAME_TO").value        = bAgnQuery.FieldByname("AGNLASTNAME_TO").value
				StoredProc.ParamByName("AGNFAMILYNAME_FR").value= bAgnQuery.FieldByname("AGNFAMILYNAME_FR").value
				StoredProc.ParamByName("AGNFIRSTNAME_FR").value        = bAgnQuery.FieldByname("AGNFIRSTNAME_FR").value
				StoredProc.ParamByName("AGNLASTNAME_FR").value        = bAgnQuery.FieldByname("AGNLASTNAME_FR").value
				StoredProc.ParamByName("AGNFAMILYNAME_AC").value= bAgnQuery.FieldByname("AGNFAMILYNAME_AC").value
				StoredProc.ParamByName("AGNFIRSTNAME_AC").value        = bAgnQuery.FieldByname("AGNFIRSTNAME_AC").value
				StoredProc.ParamByName("AGNLASTNAME_AC").value        = bAgnQuery.FieldByname("AGNLASTNAME_AC").value
				StoredProc.ParamByName("AGNFAMILYNAME_ABL").value= bAgnQuery.FieldByname("AGNFAMILYNAME_ABL").value
				StoredProc.ParamByName("AGNFIRSTNAME_ABL").value= bAgnQuery.FieldByname("AGNFIRSTNAME_ABL").value
				StoredProc.ParamByName("AGNLASTNAME_ABL").value        = bAgnQuery.FieldByname("AGNLASTNAME_ABL").value
				StoredProc.ParamByName("EMP").value                                = 0
				StoredProc.ParamByName("EMPPOST").value                        = bAgnQuery.FieldByname("AGNLASTNAME_ABL").value
				StoredProc.ParamByName("EMPPOST_FROM").value        = bAgnQuery.FieldByname("EMPPOST_FROM").value
				StoredProc.ParamByName("EMPPOST_TO").value                = bAgnQuery.FieldByname("EMPPOST_TO").value
				StoredProc.ParamByName("EMPPOST_AC").value                = bAgnQuery.FieldByname("EMPPOST_AC").value
				StoredProc.ParamByName("EMPPOST_ABL").value                = bAgnQuery.FieldByname("EMPPOST_ABL").value
				StoredProc.ParamByName("AGNBURN").value                        = bAgnQuery.FieldByname("AGNBURN").value
				StoredProc.ParamByName("PHONE").value                        = bAgnQuery.FieldByname("PHONE").value
				StoredProc.ParamByName("PHONE2").value                        = bAgnQuery.FieldByname("PHONE2").value
				StoredProc.ParamByName("FAX").value                                = bAgnQuery.FieldByname("FAX").value
				StoredProc.ParamByName("TELEX").value                        = bAgnQuery.FieldByname("TELEX").value
				StoredProc.ParamByName("MAIL").value                        = bAgnQuery.FieldByname("MAIL").value
				StoredProc.ParamByName("IMAGE").value                        = bAgnQuery.FieldByname("IMAGE").value
				StoredProc.ParamByName("dDISCDATE").value                = bAgnQuery.FieldByname("DISCDATE").value
				StoredProc.ParamByName("AGN_COMMENT").value                = bAgnQuery.FieldByname("AGN_COMMENT").value
				StoredProc.ParamByName("nSEX").value                        = 0
				StoredProc.ParamByName("sPENSION_NBR").value        = bAgnQuery.FieldByname("PENSION_NBR").value
				StoredProc.ParamByName("sMEDPOLICY_SER").value        = bAgnQuery.FieldByname("MEDPOLICY_SER").value
				StoredProc.ParamByName("sMEDPOLICY_NUMB").value        = bAgnQuery.FieldByname("MEDPOLICY_NUMB").value
				StoredProc.ParamByName("sPROPFORM").value                = propform
				StoredProc.ParamByName("sREASON_CODE").value        = bAgnQuery.FieldByname("REASON_CODE").value
				StoredProc.ParamByName("nRESIDENT_SIGN").value        = 0
				StoredProc.ParamByName("sTAXPSTATUS").value                = taxpstatus
				StoredProc.ParamByName("sOGRN").value                        = bAgnQuery.FieldByname("OGRN").value
				StoredProc.ParamByName("sPRFMLSTS").value                = prfmlsts
				StoredProc.ParamByName("sPRNATION").value                = prnation
				StoredProc.ParamByName("sCITIZENSHIP").value        = citizenship
				'StoredProc.ParamByName("CITIZENOKIN").value                = bAgnQuery.FieldByname("CITIZENOKIN").value
				StoredProc.ParamByName("ADDR_BURN").value                = bAgnQuery.FieldByname("ADDR_BURN").value
				StoredProc.ParamByName("sPRMLREL").value                = prmlrel
				StoredProc.ParamByName("sOKATO").value                        = OKATO
				StoredProc.ParamByName("sPFR_NAME").value                = bAgnQuery.FieldByname("PFR_NAME").value
				StoredProc.ParamByName("dPFR_FILL_DATE").value        = bAgnQuery.FieldByname("PFR_FILL_DATE").value
				StoredProc.ParamByName("dPFR_REG_DATE").value        = bAgnQuery.FieldByname("PFR_REG_DATE").value
				StoredProc.ParamByName("sPFR_REG_NUMB").value        = bAgnQuery.FieldByname("PFR_REG_NUMB").value
				StoredProc.ParamByName("sFULLNAME").value                = bAgnQuery.FieldByname("FULLNAME").value
				StoredProc.ParamByName("sOKFS").value                        = bAgnQuery.FieldByname("OKFS").value
				StoredProc.ParamByName("sOKOPF").value                        = bAgnQuery.FieldByname("OKOPF").value
				StoredProc.ParamByName("sTFOMS").value                        = bAgnQuery.FieldByname("TFOMS").value
				StoredProc.ParamByName("sFSS_REG_NUMB").value        = bAgnQuery.FieldByname("FSS_REG_NUMB").value
				StoredProc.ParamByName("sFSS_SUBCODE").value        = bAgnQuery.FieldByname("FSS_SUBCODE").value
				StoredProc.ParamByName("nCOEFFIC").value                = 0
				StoredProc.ParamByName("dAGNDEATH").value                = bAgnQuery.FieldByname("AGNDEATH").value
				StoredProc.ParamByName("sOKTMO").value                        = OKTMO
				StoredProc.ParamByName("sINN_CITIZENSHIP").value= bAgnQuery.FieldByname("INN_CITIZENSHIP").value
				StoredProc.ParamByName("nIND_BUSINESSMAN").value=0
				StoredProc.ExecProc

				Query.SQL.Text = "select * from AGNBANKS where RN='"&agnbank_rn&"'"
				Query.Open
				'ОБНОВИМ ЗАПИСЬ ОБ ЭЛЕМЕНТЕ СЛОВАРЯ "БАНКОВСКИЕ УЧРЕЖДЕНИЯ"
				StoredProc.StoredProcName="P_AGNBANKS_UPDATE"
				StoredProc.ParamByName("nCOMPANY").value                = 42903
				StoredProc.ParamByName("nRN").value                                = agnbank_rn
				StoredProc.ParamByName("sBANKFCODEACC").value        = nodeNode.selectSingleNode("БИК").text
				StoredProc.ParamByName("sBANKACC").value                = nodeNode.selectSingleNode("КоррСчет").text
				StoredProc.ParamByName("sCODE").value                        = "БАНК_"&nodeNode.selectSingleNode("БИК").text
				StoredProc.ParamByName("sSWIFT").value                        = Query.FieldByname("SWIFT").value
				StoredProc.ParamByName("sMEMBER_CODE").value        = Query.FieldByname("MEMBER_CODE").value
				StoredProc.ParamByName("sMEMBER_NAME").value        = Query.FieldByname("MEMBER_NAME").value
				StoredProc.ParamByName("sMEMBER_REG").value                = Query.FieldByname("MEMBER_REG").value
				StoredProc.ExecProc
			end if
		next
		MyFile.Write("INFO "&now()&vbTab&" Банки загружены. Всего обработано: "& object_counter & vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" В текущей выгрузке на найдено информации о банках."&vbNewLine)
	end if

	'НАХОДИМ УЗЕЛ "СЧЕТА"
	Set Accounts = xmlParser.selectNodes("//БанковскиеСчета/Строки")
	If Accounts.length > 0 then
		'ПЕРЕБИРАЕМ СПИСОК СЧЕТОВ
		object_counter=0
		For Each nodeNode In Accounts
			account_rn                = NULL
			AccQueryIsEmpty        = NULL
			agnlist_rn                = NULL
			agnlist_name        = NULL
			agnbank_mnemo        = NULL
			lastcode                = NULL
			counter                        = NULL
			newRN                        = NULL

			object_counter=object_counter+1

			'ПОЛУЧИМ RN КОНТРАГЕНТА ПО КОДУ SAP
			Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&RTrim(nodeNode.selectSingleNode("Владелец").text)&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - код свойства "Код SAP"
			Query.Open
			agnlist_rn=Query.FieldByname("UNIT_RN").value
			Query.Close

			If not agnlist_rn=0 then
				If not nodeNode.selectSingleNode("Банк").text=""  then
					'ПОЛУЧИМ ДАННЫЕ БАНКА
					Query.SQL.Text = "select SNAME, SBANKACC, SCODE from v_AGNBANKS where SCODE='БАНК_"&nodeNode.selectSingleNode("Банк").text&"'"
					Query.Open
					agnbank_mnemo = Query.FieldByname("SCODE").value

					'ПОЛУЧИМ НАИМЕНОВАНИЕ КОНТРАГЕНТА ДЛЯ НАИМЕНОВАНИЯ СЧЕТА
					Query.SQL.Text= "select AGNNAME from AGNLIST where RN='"&agnlist_rn&"'"
					Query.Open
					agnlist_name = Query.FieldByname("AGNNAME").value
					Query.Close

					Query.SQL.Text="select * from v_agnacc where AGNACC='"&nodeNode.selectSingleNode("НомерСчета").text&"' and AGNRN='"&agnlist_rn&"' and SBANKCODEACC='"&agnbank_mnemo&"' order by STRCODE"
					Query.Open
					Query.Last

					If Query.IsEmpty then
						MyFile.Write("INFO "&now()&vbTab&" В Парус не найден банковский счет с номером "&nodeNode.selectSingleNode("НомерСчета").text&"  - создаю новый счет."&vbNewLine)

						'НАЙДЕМ ПАРАМЕТР "КОД СТРОКИ"
						REM StoredProc.StoredProcName="FIND_AGNACC_LASTCODE"
						REM StoredProc.ParamByName("COMPANY").value        = 42903                                        'код подразделения
						REM StoredProc.ParamByName("AGNRN").value        = agnlist_rn                                        'код контрагента
						REM StoredProc.ExecProc
						Query.SQL.Text="select max(strcode) from AGNACC where agnrn = '"&agnlist_rn&"'"                    'новое 05.05.2017
						Query.Open
						If not Query.FieldByname("max(strcode)").value = "" then
							lastcode = Query.FieldByname("max(strcode)").value
							lastcode=lastcode+1
							lastcode=CStr(lastcode)
							counter = 4-len(lastcode)
							do while counter > 0
								lastcode="0"&lastcode
								counter=counter-1
							loop
						else
							lastcode = "0001"
						end if

						'СОЗДАЕМ ЗАПИСЬ О НОВОМ БАНКОВСКОМ СЧЕТЕ
						StoredProc.StoredProcName="P_AGNACC_INSERT"
						StoredProc.ParamByName("nCOMPANY").value                = 42903                                        'код подразделения
						StoredProc.ParamByName("nPRN").value                        = agnlist_rn                                        'код контрагента
						StoredProc.ParamByName("sSTRCODE").value                = lastcode
						StoredProc.ParamByName("SAGNACC").value                        = nodeNode.selectSingleNode("НомерСчета").text
						StoredProc.ParamByName("sAGNNAMEACC").value                = agnlist_name
						StoredProc.ParamByName("SAGNBANKS").value                = agnbank_mnemo
						StoredProc.ParamByName("NACCESS_FLAG").value        = 1
						StoredProc.ExecProc
						newRN = StoredProc.ParamByName("nRN").value

						REM 'ВНЕСЕМ КОД 1С В СВОЙСТВА
						REM 'StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'процедурка для записи доп свойства
						REM 'StoredProc.ParamByName("PROPERTY").value="СчетКод1С"
						REM 'StoredProc.ParamByName("UNITCODE").value="ContragentsBankAttrs"
						REM 'StoredProc.ParamByName("RN_SOTR").value=newRN
						REM 'StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("Ссылка").text)
						REM 'StoredProc.ParamByName("NUM_VAL").value=NULL
						REM 'StoredProc.ExecProc
						Query.Close
					else
						MyFile.Write("INFO "&now()&vbTab&" В Парус найден банковский счет с номером "&nodeNode.selectSingleNode("НомерСчета").text&"  - обновляю существующий счет."&vbNewLine)

						account_rn = Query.FieldByname("RN").value

						'НАЙДЕМ НОМЕР СТРОКИ ДЛЯ ТЕКУЩЕЙ ЗАПИСИ О БАНКОВСКОМ СЧЕТЕ
						Query.SQL.Text = "select * from AGNACC where RN = '"&account_rn&"'"
						Query.Open

						STRCODE                        = Query.FieldByname("STRCODE").value
						BANKNAMEACC                = Query.FieldByname("BANKNAMEACC").value
						BANKFCODEACC        = Query.FieldByname("BANKFCODEACC").value
						BANKACC                        = Query.FieldByname("BANKACC").value
						BANKCITYACC                = Query.FieldByname("BANKCITYACC").value
						OPEN_DATE                = Query.FieldByname("OPEN_DATE").value
						CLOSE_DATE                = Query.FieldByname("CLOSE_DATE").value
						COUNTRY_CODE        = Query.FieldByname("COUNTRY_CODE").value
						SWIFT                        = Query.FieldByname("SWIFT").value
						REGION                        = Query.FieldByname("REGION").value
						DISTRICT                = Query.FieldByname("DISTRICT").value
						BANKACC_TYPE        = Query.FieldByname("BANKACC_TYPE").value
						sCURRENCY                = Query.FieldByname("CURRENCY").value
						CORR_AGNACC                = Query.FieldByname("CORR_AGNACC").value
						CARDNUMB                = Query.FieldByname("CARDNUMB").value
						AGNTREAS                = Query.FieldByname("AGNTREAS").value
						REAS_AGNACC                = Query.FieldByname("TREAS_AGNACC").value
						INTERMEDIARY        = Query.FieldByname("INTERMEDIARY").value
						INTERMED_ACC        = Query.FieldByname("INTERMED_ACC").value

						'УСТАНОВИМ В NULL ВСЕ ЗНАЧЕНИЯ 0
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

						'ОБНОВИМ ЗАПИСЬ О БАНКОВСКОМ СЧЕТЕ
						StoredProc.StoredProcName="P_AGNACC_UPDATE"
						StoredProc.ParamByName("nCOMPANY").value                = 42903
						StoredProc.ParamByName("nRN").value                                = account_rn
						StoredProc.ParamByName("sSTRCODE").value                = STRCODE
						StoredProc.ParamByName("SAGNACC").value                        = nodeNode.selectSingleNode("НомерСчета").text
						StoredProc.ParamByName("sAGNNAMEACC").value                = agnlist_name
						StoredProc.ParamByName("sBANKNAMEACC").value        = BANKNAMEACC
						StoredProc.ParamByName("sBANKFCODEACC").value        = BANKFCODEACC
						StoredProc.ParamByName("sBANKACC").value                = BANKACC
						StoredProc.ParamByName("sBANKCITYACC").value        = BANKCITYACC
						StoredProc.ParamByName("sAGNBANKS").value                = agnbank_mnemo
						StoredProc.ParamByName("dOPEN_DATE").value                = OPEN_DATE
						StoredProc.ParamByName("dCLOSE_DATE").value                = CLOSE_DATE
						StoredProc.ParamByName("sCOUNTRY_CODE").value        = COUNTRY_CODE
						StoredProc.ParamByName("nACCESS_FLAG").value        = 1
						StoredProc.ParamByName("sSWIFT").value                        = SWIFT
						StoredProc.ParamByName("sREGION").value                        = REGION
						StoredProc.ParamByName("sDISTRICT").value                = DISTRICT
						StoredProc.ParamByName("sBANKACC_TYPE").value        = BANKACC_TYPE
						StoredProc.ParamByName("sCURRENCY").value                = sCURRENCY
						StoredProc.ParamByName("sCORR_AGNACC").value        = CORR_AGNACC
						StoredProc.ParamByName("sCARDNUMB").value                = CARDNUMB
						StoredProc.ParamByName("sAGNTREAS").value                = AGNTREAS
						StoredProc.ParamByName("sTREAS_AGNACC").value        = REAS_AGNACC
						StoredProc.ParamByName("sINTERMEDIARY").value        = INTERMEDIARY
						StoredProc.ParamByName("sINTERMED_ACC").value        = INTERMED_ACC
						StoredProc.ExecProc
					end if
				else
					agnbank_mnemo = NULL
				end if
			else
				If nodeNode.selectSingleNode("ЭтоСчетОрганизации").text = "true" or nodeNode.selectSingleNode("Примечание").text ="Нижнекамская ТЭЦ ООО"  then
					MyFile.Write("INFO "&now()&vbTab&" Cчет с номером "&nodeNode.selectSingleNode("НомерСчета").text&" принадлежит ООО <Нижнекамская ТЭЦ> - пропускаю счет."&vbNewLine)
				elseIf nodeNode.selectSingleNode("Владелец").text = "" then
					MyFile.Write("INFO "&now()&vbTab&"У счета с номером "&nodeNode.selectSingleNode("НомерСчета").text&" не указан владелец - пропускаю счет."&vbNewLine)
				elseIf nodeNode.selectSingleNode("ВидСчета").text = "Депозитный" then
					MyFile.Write("INFO "&now()&vbTab&" Cчет с номером "&nodeNode.selectSingleNode("НомерСчета").text&" является депозитным - пропускаю счет."&vbNewLine)
				else
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" В Парус не найден контрагент с кодом SAP "&nodeNode.selectSingleNode("Владелец").text&" которому принадлежит счет с номером "&nodeNode.selectSingleNode("НомерСчета").text&"  - не могу создать/обновить счет."&vbNewLine)
				end if
			end if
		next
		MyFile.Write("INFO "&now()&vbTab&" Банковские счета загружены. Всего обработано: "& object_counter & vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" В текущей выгрузке на найдено информации о банковских счетах."&vbNewLine)
	end if

	'НАХОДИМ УЗЕЛ "ДОГОВОРА"'
	Set Contracts = xmlParser.selectNodes("//ДоговорыКонтрагентов/Строки")
	If Contracts.length > 0 then
		object_counter = 0
		For Each nodeNode In Contracts
			comment             	= NULL
			comment_array          	= NULL
			contract_complex_number	= NULL
			contract_number_array  	= NULL
			doc_type_RN				= NULL
			old_RN                 	= NULL
			agn_abbr              	= NULL
			agn_rn                	= NULL
			orgaccbik           	= NULL
			orgacc                	= NULL
			jur_strcode          	= NULL
			agnaccbik             	= NULL
			agnacc                 	= NULL
			agn_strcode          	= NULL
			INOUT_SIGN              = NULL
			ext_agreement      		= 0
			PRN        				= NULL
			executive           	= NULL
			subdiv               	= NULL
			scurrency        		= NULL
			newContract     		= NULL
			doc_numb         		= NULL
			warning            		= NULL
			DogovorZaima       		= NULL

			object_counter = object_counter+1

			Query.SQL.Text = "select AGNABBR from agnlist where agnname like upper('%"&nodeNode.selectSingleNode("ОтветственныйИсполнитель").text&"%') and EMP=1 order by RN DESC"
			Query.Open
			If Query.IsEmpty and not InStr(nodeNode.selectSingleNode("ТипДоговора").text, "депозит")=0 then
				MyFile.Write("INFO "&now()&vbTab&" Договор "&comment&" ("&trim(nodeNode.selectSingleNode("Ссылка").text)&") - депозитный, пропускаю договор."&vbNewLine)
			else
				If Query.IsEmpty and not InStr(nodeNode.selectSingleNode("ТипДоговора").text, "займ")=0 then
					DogovorZaima = True
				end if

				'НАЙДЕМ ДОГОВОР ПО ССЫЛКЕ 1С
				Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value like '%"&trim(nodeNode.selectSingleNode("Ссылка").text)&"%' and docs_prop_rn='104582667' and unitcode='Contracts'"        ' 87456099 - код свойства "Код 1С"
				Query.Open
				If Query.IsEmpty then
'''''''''''''''''''''УДАЛИТЬ БЛОК ПРИСВОЕНИЯ НОМЕРА ПО КОММЕНТАРИЮ, С 01.01.2018 ДЕЙСТВУЕТ АВТОМАТИЧЕСКАЯ НУМЕРАЦИЯ
					'ПОДГОТОВИМ ЗАГЛУШКИ НА СЛУЧАЙ, ЕСЛИ НЕ УДАСТСЯ РАЗОБРАТЬ КОММЕНТАРИЙ
					doc_type                = "0000"
					doc_pref                = date&"/"&timer
					'ПОЛУЧИМ СВОБОДНЫЙ ПОРЯДКОВЫЙ НОМЕР
					StoredProc.StoredProcName="P_CONTRACTS_GETNEXTNUMB"
					StoredProc.ParamByName("NCOMPANY").value=42903
					StoredProc.ParamByName("SJUR_PERS").value="НК ТЭЦ"
					StoredProc.ParamByName("DDOC_DATE").value=ConvDate(nodeNode.selectSingleNode("Дата").text)
					StoredProc.ParamByName("SDOC_TYPE").value=doc_type
					StoredProc.ParamByName("SDOC_PREF").value=doc_pref
					StoredProc.ExecProc
					doc_numb                        = StoredProc.ParamByName("SDOC_NUMB").value
					CONTRACT_NEXTNUMB        = StoredProc.ParamByName("SDOC_NUMB").value

					'ЕСЛИ КОММЕНТАРИЙ ЗАПОЛНЕН - ПОПРОБУЕМ ВЫТАЩИТЬ ИЗ НЕГО ДАННЫЕ ПО РАНЕЕ ЗАГРУЖЕННОМУ ДОГОВОРУ					
					comment = nodeNode.selectSingleNode("Комментарий").text
					If not len(comment)=0 then
						'РАСПИЛИМ КОММЕНТАРИЙ НА СОСТАВНЫЕ ЧАСТИ: ТИП, ПРЕФИКС И НОМЕР, ПРОВЕРЯЯ КОЛИЧЕСТВО СОСТАВНЫХ ЧАСТЕЙ						
						comment_array                        = Split(LTrim(comment), ",")
						if UBound(comment_array)>0 then
							contract_complex_number        = RTrim(LTrim(comment_array(1)))
							contract_number_array        = Split(contract_complex_number, "-")
							if UBound(contract_number_array)>0 then
								doc_type                = comment_array(0)
								doc_pref                = contract_number_array(0)
								doc_numb                = contract_number_array(1)

								'НАЙДЕМ КОД ТИПА ДОКУМЕНТА
								Query.SQL.Text = "select RN from DOCTYPES where DOCCODE='"&doc_type&"'"
								Query.Open
								doc_type_RN        = Query.FieldByname("RN").value
								Query.Close

								'НАЙДЕМ ДОГОВОР ПО ТИПУ, ПРЕФИКСУ И НОМЕРУ
								If not doc_numb="" then
									Query.SQL.Text = "select RN from contracts where doc_type='"&doc_type_RN&"' and DOC_PREF like '%"&doc_pref&"' and DOC_NUMB like '%"&doc_numb&"'"
									Query.Open
									If not Query.IsEmpty and (nodeNode.selectSingleNode("БазовыйДоговор").text="00000000-0000-0000-0000-000000000000" or len(nodeNode.selectSingleNode("БазовыйДоговор").text)=0) then
										newContract = False
										old_RN = Query.FieldByname("RN").value                'ЗАПОМНИМ УИН ДЛЯ СУЩЕСТВУЮЩЕГО ДОГОВОРА
									else
										newContract = True
									end if
									Query.Close
								else
									doc_numb = CONTRACT_NEXTNUMB
									newContract = True
								end if
							else
								newContract = True
							end if
						else
							newContract = True
						end if
					else
						newContract = True
					end If
''''''''''''''''''''УДАЛИТЬ БЛОК ПРИСВОЕНИЯ НОМЕРА ПО КОММЕНТАРИЮ, С 01.01.2018 ДЕЙСТВУЕТ АВТОМАТИЧЕСКАЯ НУМЕРАЦИЯ					
''''''''''''''''''''newContract = True
				else
					newContract = False
					old_RN = Query.FieldByname("UNIT_RN").value                'ЗАПОМНИМ УИН ДЛЯ СУЩЕСТВУЮЩЕГО ДОГОВОРА
				end if

				'ЕСЛИ ДОГОВОР НЕ НАЙДЕН ПО КОДУ 1С ИЛИ СТРОКЕ КОММЕНТАРИЯ - СОЗДАДИМ НОВЫЙ
				If newContract then
					MyFile.Write("INFO "&now()&vbTab&" В Парус не найден договор "&comment&" ("&trim(nodeNode.selectSingleNode("Ссылка").text)&"), либо не удалось правильно разобрать поле Комментарий - пробую создать новый договор."&vbNewLine)

					'ПОЛУЧИМ МНЕМОКОД КОНТРАГЕНТА ПО КОДУ SAP
					Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&RTrim(nodeNode.selectSingleNode("Владелец").text)&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - код свойства "Код SAP" (1ИЗМЕНИТЬ ПОСЛЕ ПЕРЕНОСА В ПРОДУКТИВ)
					Query.Open
					Query.SQL.Text = "select AGNABBR, RN, AGNNAME from AGNLIST where RN='" & Query.FieldByname("UNIT_RN").value & "'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value
					agn_rn = Query.FieldByname("RN").value
					agn_name = Query.FieldByname("AGNNAME").value
					Query.Close

					'ПОЛУЧИМ НОМЕР СТРОКИ БАНКОВСКОГО СЧЕТА ОРГАНИЗАЦИИ ЧЕРЕЗ НОМЕР ЭТОГО СЧЕТА
					If not nodeNode.selectSingleNode("СчетОрганизации").text = "" then
						orgaccbik                = Trim(nodeNode.selectSingleNode("СчетОрганизацииБИК").text)
						orgacc                        = Trim(nodeNode.selectSingleNode("СчетОрганизации").text)
						Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & orgaccbik & "'"
						Query.Open
						If not Query.IsEmpty then
							Query.SQL.Text        = "select STRCODE from AGNACC where AGNACC='"& orgacc &"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='5805775'"
							Query.Open
							jur_strcode                = Query.FieldByname("STRCODE").value
							Query.Close
						else
							MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Для банковского счета с номером "&nodeNode.selectSingleNode("СчетОрганизации").text&" в Парус не найден БИК "&nodeNode.selectSingleNode("СчетОрганизацииБИК").text&" - делаю останов для отладки."&vbNewLine)
						end if
					else
						jur_strcode                = NULL
					end if

					'ПОЛУЧИМ НОМЕР СТРОКИ БАНКОВСКОГО СЧЕТА КОНТРАГЕНТА ЧЕРЕЗ НОМЕР ЭТОГО СЧЕТА
					if not nodeNode.selectSingleNode("СчетКонтрагента").text="" then
						agnaccbik                = Trim(nodeNode.selectSingleNode("СчетКонтрагентаБИК").text)
						agnacc                        = Trim(nodeNode.selectSingleNode("СчетКонтрагента").text)
						Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & agnaccbik & "'"
						Query.Open
						If not Query.IsEmpty then
							Query.SQL.Text = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='"&agn_rn&"'"
							Query.Open
							If not Query.IsEmpty then
								agn_strcode = Query.FieldByname("STRCODE").value
								note = nodeNode.selectSingleNode("ПредметДоговора").text
							else
								MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Банковский счет с номером "&nodeNode.selectSingleNode("СчетКонтрагента").text&" не принадлежит контрагенту "&agn_abbr&"  - в договоре <<"&comment&">> будет указан случайный счет контрагента"&vbNewLine)
								Query.SQL.Text = "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
								Query.Open
								agn_strcode = Query.FieldByname("STRCODE").value
								note = "В ДОГОВОРЕ УКАЗАН СЛУЧАЙНЫЙ БАНКОВСКИЙ СЧЕТ КОНТРАГЕНТА - ВЫБЕРИТЕ ПРАВИЛЬНЫЙ СЧЕТ!!!"
							end if
							Query.Close
						else
							MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Для банковского счета с номером "&nodeNode.selectSingleNode("СчетКонтрагента").text&" в Парус не найден БИК "&nodeNode.selectSingleNode("СчетКонтрагентаБИК").text&" - делаю останов для отладки."&vbNewLine)
						end if
					else
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" В Парус не найден банковский счет контрагента "&agn_abbr&" - в договоре <<"&comment&">> будет указан случайный счет контрагента"&vbNewLine)
						Query.SQL.Text = "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
						Query.Open
						if not Query.IsEmpty then
							agn_strcode = Query.FieldByname("STRCODE").value
						else
							'СОЗДАЕМ ЗАПИСЬ О НОВОМ БАНКОВСКОМ СЧЕТЕ
							StoredProc.StoredProcName="P_AGNACC_INSERT"
							StoredProc.ParamByName("nCOMPANY").value                = 42903
							StoredProc.ParamByName("nPRN").value                        = agn_rn
							StoredProc.ParamByName("sSTRCODE").value                = "0001"
							StoredProc.ParamByName("SAGNACC").value                        = NULL 'nodeNode.selectSingleNode("НомерСчета").text
							StoredProc.ParamByName("sAGNNAMEACC").value                = agn_name
							StoredProc.ParamByName("SAGNBANKS").value                = NULL
							StoredProc.ParamByName("NACCESS_FLAG").value        = 1
							StoredProc.ExecProc
							agn_strcode = "0001"
						end if
						note = "В ДОГОВОРЕ УКАЗАН СЛУЧАЙНЫЙ БАНКОВСКИЙ СЧЕТ КОНТРАГЕНТА - ВЫБЕРИТЕ ПРАВИЛЬНЫЙ СЧЕТ!!!"
					end if

					'ЗАПОЛНИМ НЕКОТОРЫЕ ФЛАГИ
					if len(nodeNode.selectSingleNode("ВнешРегНомер").text)>0 then
						INOUT_SIGN = 0
					else
						INOUT_SIGN = 1
					end if

					'ПОЛУЧИМ КОД ПОДРАЗДЕЛЕНИЯ ПО ЦЕПОЧКЕ: "ФИО ИСПОЛНИТЕЛЯ -> УИН КОНТРАГЕНТА -> УИН СОТРУДНИКА -> ЗАПИСЬ О ТЕКУЩЕЙ ДОЛЖНОСТИ -> УИН ПОДРАЗДЕЛЕНИЯ -> КОД ПОДРАЗДЕЛЕНИЯ"
					if nodeNode.selectSingleNode("ОтветственныйИсполнитель").text="Гатина Гузель Илдаровна" then
						executive = "0001 ГАТИНА Г.И."                'ГАТИНА Г.И. - ПРОФКОМ, ЕЕ НЕТ СРЕДИ СОТРУДНИКОВ ТЭЦ - НАЗНАЧЕМ ПОДРАЗДЕЛЕНИЕ "ОТДЕЛ КАДРОВ"
						subdiv = "НкТЭЦ.13.15"
					elseIf DogovorZaima then
						executive = "6062 ТУХВАТУЛЛИНА"                'ТУХВАТУЛЛИНА М.Ф. ведет договора займов, НАЗНАЧЕМ ПОДРАЗДЕЛЕНИЕ "финансовый отдел"
						subdiv = "НкТЭЦ.13.18"
					elseif nodeNode.selectSingleNode("ОтветственныйИсполнитель").text="Мартынова Оксана Николаевна" then
						executive = "5761 МАРТЫНОВА О.Н."                'Мартынова О.Н. числится в СПЛ, но фактически работает в ОПК
						subdiv = "НкТЭЦ.13.09"
					else
						Query.SQL.Text = "select a.AGNABBR, a.AGNNAME, a.RN, b.code from agnlist a, CLNPERSONS b where a.rn=b.pers_agent and agnname like upper('%"&nodeNode.selectSingleNode("ОтветственныйИсполнитель").text&"%') and not b.crn=2503442 and EMP=1 and DISMISS_DATE is NULL order by RN DESC"        'ОТСЕИМ НЕ СОТРУДНИКОВ И СОТРУДНИКОВ ИЗ ПАПКИ УВОЛЕННЫЕ
						Query.Open
						executive                = Query.FieldByname("AGNABBR").value
						Query.SQL.Text = "select RN from CLNPERSONS where PERS_AGENT='"&Query.FieldByname("RN").value&"'"
						Query.Open
						Query.SQL.Text = "select DEPTRN from CLNPSPFM where persrn='"&Query.FieldByname("RN").value&"' and endeng is null"
						Query.Open
						Query.SQL.Text = "select CODE from INS_DEPARTMENT where rn='"&Query.FieldByname("DEPTRN").value&"'"
						Query.Open
						subdiv = Query.FieldByname("CODE").value
						Query.Close
					end if
										
					'ПОЛУЧИМ НАИМЕНОВАНИЕ ВАЛЮТЫ ПО ЕЕ КОДУ
					Query.SQL.Text = "select INTCODE from curnames where curcode='"&nodeNode.selectSingleNode("ВалютаВзаиморасчетов").text&"'"
					Query.Open
					If Query.IsEmpty then
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Валюта с кодом "&nodeNode.selectSingleNode("ВалютаВзаиморасчетов").text&" из договора <<"&comment&">> не найдена. Используется значение по-умолчанию - RUR"&vbNewLine)
						scurrency = "RUR"
					else
						scurrency = nodeNode.selectSingleNode("ВалютаВзаиморасчетов").text
					end if
					Query.Close

					If nodeNode.selectSingleNode("СрокДействияПо").text="0001-01-01" then
						endDate = "01.01.0001"        'NULL
					Else
						endDate = ConvDate(nodeNode.selectSingleNode("СрокДействияПо").text)
					end if

					If agn_abbr="" then
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" В Парус не найден контрагент с кодом "&nodeNode.selectSingleNode("Владелец").text&" из договора "&comment&" - пропускаю договор."&vbNewLine)
					else

						'ПОЛУЧИМ ТИП СУММЫ И ЛСЧЕТА ДЛЯ ЭТАПА
						If nodeNode.selectSingleNode("ВидДоговора").text = "С поставщиком" then
							sumType = 1
							acc_kind = 0
						elseif nodeNode.selectSingleNode("ВидДоговора").text = "С покупателем" then
							sumType = 2
							acc_kind = 1
						else
							sumType = 0
							acc_kind = 0
						end if

						'ОПРЕДЕЛИМ - ДОПСОГЛАШЕНИЕ ИЛИ НОВЫЙ ДОГОВОР
						if nodeNode.selectSingleNode("БазовыйДоговор").text="00000000-0000-0000-0000-000000000000" or len(nodeNode.selectSingleNode("БазовыйДоговор").text)=0 then
							'СФОРМИРУЕМ НОМЕР ДОГОВОРА
							'номер договора
							doc_date_array = Split(Trim(nodeNode.selectSingleNode("СрокДействияС").text), "-")
							doc_date_year = doc_date_array(0)
							If not CInt(doc_date_year) < 2018 then
								Query.SQL.Text = "select rn, doc_pref, doc_numb from contracts where begin_date between '01-янв-"&doc_date_year&"' and '31-дек-"&doc_date_year&"' order by doc_numb desc"
								Query.Open
								If not Query.IsEmpty then
									do while not Query.EOF
										If not IsNumeric(Query.FieldByname("doc_numb").value) then
											Query.next
										else
											doc_numb_last = CInt(Query.FieldByname("doc_numb").value)
											Query.last
										end if
									loop
								else
									MyFile.Write("	ERROR "&now()&vbTab&" Договору с GUID ("&trim(nodeNode.selectSingleNode("Ссылка").text)&") НЕ удалось автоматически присвоить номер."&vbNewLine)
								end if					
								doc_numb_last = doc_numb_last+1
								doc_numb_last=CStr(doc_numb_last)
								counter = 4-len(doc_numb_last)
								do while counter > 0
									doc_numb_last="0"&doc_numb_last
									counter=counter-1
								loop
								doc_numb = doc_numb_last
								'префикс договора
								doc_pref = doc_date_year&"/"&GetContractSubdivPref(subdiv)
								note = note&" ("&doc_pref&"-"&doc_numb&")" 
								MyFile.Write("	INFO "&now()&vbTab&" Договору с GUID ("&trim(nodeNode.selectSingleNode("Ссылка").text)&") автоматически присвоен номер: "&doc_pref&"-"&doc_numb&"."&vbNewLine)
							end if
							
							sbuf="Создается новый договор №" & doc_pref & "-" & doc_numb & ". Продолжить загрузку или прервать и начать отладку (Отмена)?" & chr(13) & chr(13)
							Desc=MsgBox(sbuf, vbOKCancel)
							If Desc = 2 then
								Wscript.Echo
							end if
						
							'СОЗДАЕМ ЗАПИСЬ О НОВОМ ДОГОВОРЕ
							StoredProc.StoredProcName="P_CONTRACTS_INSERT"
							StoredProc.ParamByName("nCOMPANY").value		= 42903
							StoredProc.ParamByName("nCRN").value			= 104583519                                        'код каталога
							StoredProc.ParamByName("nPRN").value			= PRN                                        'код основного договора
							StoredProc.ParamByName("SJUR_PERS").value		= "НК ТЭЦ"
							StoredProc.ParamByName("SJUR_ACC").value		= jur_strcode
							StoredProc.ParamByName("SDOC_TYPE").value		= doc_type
							StoredProc.ParamByName("SDOC_PREF").value		= doc_pref
							StoredProc.ParamByName("SDOC_NUMB").value		= doc_numb
							StoredProc.ParamByName("DDOC_DATE").value		= ConvDate(nodeNode.selectSingleNode("Дата").text)
							StoredProc.ParamByName("SEXT_NUMBER").value		= nodeNode.selectSingleNode("ВнешРегНомер").text
							StoredProc.ParamByName("NINOUT_SIGN").value		= INOUT_SIGN        'булево Входящий, Истина=0, Ложь=1
							StoredProc.ParamByName("NFALSE_DOC").value		= 0                                'булево Условный, Ложь=0, Истина=1
							StoredProc.ParamByName("NEXT_AGREEMENT").value	= ext_agreement        'булево Допсоглашение, Ложь=0, Истина=1
							StoredProc.ParamByName("SAGENT").value			= agn_abbr
							StoredProc.ParamByName("SAGNACC").value			= agn_strcode
							StoredProc.ParamByName("SEXECUTIVE").value		= executive
							StoredProc.ParamByName("SSUBDIVISION").value	= subdiv
							StoredProc.ParamByName("DBEGIN_DATE").value		= ConvDate(nodeNode.selectSingleNode("СрокДействияС").text)
							StoredProc.ParamByName("DEND_DATE").value		= endDate
							StoredProc.ParamByName("NSUM_TYPE").value		= 1
							StoredProc.ParamByName("NDOC_SUM").value		= 0
							StoredProc.ParamByName("NDOC_SUMTAX").value		= Replace(nodeNode.selectSingleNode("СуммаДоговора").text, ".", ",")
							StoredProc.ParamByName("NDOC_SUM_NDS").value	= 0
							StoredProc.ParamByName("NAUTOCALC_SIGN").value	= 1
							StoredProc.ParamByName("SCURRENCY").value		= scurrency
							StoredProc.ParamByName("NCURCOURS").value		= 1
							StoredProc.ParamByName("NCURBASE").value		= 1
							StoredProc.ParamByName("SSUBJECT").value		= nodeNode.selectSingleNode("ПредметДоговора").text
							StoredProc.ParamByName("SNOTE").value			= note
							StoredProc.ParamByName("NGOVDEFORD_EXEC").value	= 0
							StoredProc.ExecProc
							newRN = StoredProc.ParamByName("nRN").value

							'ЗАПИШЕМ ДАННЫЕ В СВОЙСТВА ДОКУМЕНТА
							StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'ДогКод1С
							StoredProc.ParamByName("PROPERTY").value="ДогКод1С"
							StoredProc.ParamByName("UNITCODE").value="Contracts"
							StoredProc.ParamByName("RN_SOTR").value=newRN
							StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("Ссылка").text)
							StoredProc.ParamByName("NUM_VAL").value=NULL
							StoredProc.ExecProc
							
							'проверка в ЮГ не обязательна для ООРЭМ, всем остальным флаг инициализируем в ноль.
							If subdiv="НкТЭЦ.13.16" then	'ООРЭМ
								Lawyer_check = 1
							else
								Lawyer_check = 0
							end if							
							StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'ПроверенЮГ
							StoredProc.ParamByName("PROPERTY").value="Субъект мал/ср бизн."
							StoredProc.ParamByName("UNITCODE").value="Contracts"
							StoredProc.ParamByName("RN_SOTR").value=newRN
							StoredProc.ParamByName("ST_VAL").value=NULL
							StoredProc.ParamByName("NUM_VAL").value = Lawyer_check
							StoredProc.ExecProc
							
							If subdiv="НкТЭЦ.13.16" then	'ООРЭМ
								'ЗАПИШЕМ ДАННЫЕ О ГОСЗАКУПКАХ В СВОЙСТВА ДОКУМЕНТА
								StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'Способ закупки
								StoredProc.ParamByName("PROPERTY").value="Способ закупки"
								StoredProc.ParamByName("UNITCODE").value="Contracts"
								StoredProc.ParamByName("RN_SOTR").value=newRN
								StoredProc.ParamByName("ST_VAL").value=0
								StoredProc.ParamByName("NUM_VAL").value=NULL
								StoredProc.ExecProc
								
								StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'Кол-во под. зав.
								StoredProc.ParamByName("PROPERTY").value="Кол-во под. зав."
								StoredProc.ParamByName("UNITCODE").value="Contracts"
								StoredProc.ParamByName("RN_SOTR").value=newRN
								StoredProc.ParamByName("ST_VAL").value=0
								StoredProc.ParamByName("NUM_VAL").value=NULL
								StoredProc.ExecProc
								
								StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'Кол-во нед. зав.
								StoredProc.ParamByName("PROPERTY").value="Кол-во нед. зав."
								StoredProc.ParamByName("UNITCODE").value="Contracts"
								StoredProc.ParamByName("RN_SOTR").value=newRN
								StoredProc.ParamByName("ST_VAL").value=0
								StoredProc.ParamByName("NUM_VAL").value=NULL
								StoredProc.ExecProc
								
								StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'нач. цена
								StoredProc.ParamByName("PROPERTY").value="нач. цена"
								StoredProc.ParamByName("UNITCODE").value="Contracts"
								StoredProc.ParamByName("RN_SOTR").value=newRN
								StoredProc.ParamByName("ST_VAL").value=0
								StoredProc.ParamByName("NUM_VAL").value=NULL
								StoredProc.ExecProc
								
								StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'номер закуп.проц.
								StoredProc.ParamByName("PROPERTY").value="номер закуп.проц."
								StoredProc.ParamByName("UNITCODE").value="Contracts"
								StoredProc.ParamByName("RN_SOTR").value=newRN
								StoredProc.ParamByName("ST_VAL").value=0
								StoredProc.ParamByName("NUM_VAL").value=NULL
								StoredProc.ExecProc
								
								StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'Субъект мал/ср бизн.
								StoredProc.ParamByName("PROPERTY").value="Субъект мал/ср бизн."
								StoredProc.ParamByName("UNITCODE").value="Contracts"
								StoredProc.ParamByName("RN_SOTR").value=newRN
								StoredProc.ParamByName("ST_VAL").value=NULL
								StoredProc.ParamByName("NUM_VAL").value=0
								StoredProc.ExecProc	
							end if

							'ДОБАВИМ ПЕРВЫЙ ЭТАП К ДОГОВОРУ
							StoredProc.StoredProcName="P_STAGES_INSERT"
							StoredProc.ParamByName("nCOMPANY").value                = 42903
							StoredProc.ParamByName("nPRN").value                        = newRN
							StoredProc.ParamByName("SNUMB").value                        = "1"
							StoredProc.ParamByName("NEXT_AGREEMENT").value        = 0
							StoredProc.ParamByName("NSIGN_SUM").value                = 1
							StoredProc.ParamByName("DBEGIN_DATE").value                = ConvDate(nodeNode.selectSingleNode("СрокДействияС").text)
							StoredProc.ParamByName("DEND_DATE").value                = endDate
							StoredProc.ParamByName("SJUR_ACC").value                = jur_strcode
							StoredProc.ParamByName("NSUM_TYPE").value                = sumType                                'расчет суммы
							StoredProc.ParamByName("NSTAGE_SUM").value                = 0
							StoredProc.ParamByName("NSTAGE_SUMTAX").value        = Replace(nodeNode.selectSingleNode("СуммаДоговора").text, ".", ",")
							StoredProc.ParamByName("NSTAGE_SUM_NDS").value        = 0
							StoredProc.ParamByName("NAUTOCALC_SIGN").value        = 1
							StoredProc.ParamByName("SDESCRIPTION").value        = nodeNode.selectSingleNode("ПредметДоговора").text
							StoredProc.ParamByName("SCOMMENTS").value                = nodeNode.selectSingleNode("ПредметДоговора").text
							StoredProc.ParamByName("NFACEACC_EXIST").value        = 0
							StoredProc.ParamByName("SFACEACCCRN").value                = GetFaceAccCat(subdiv)'связать по подразделению
							StoredProc.ParamByName("SAGENT").value                        = agn_abbr
							StoredProc.ParamByName("SFACEACC").value                = doc_pref&"/"&doc_numb&"/1"
							StoredProc.ParamByName("NACC_KIND").value                = acc_kind
							StoredProc.ParamByName("SEXECUTIVE").value                = executive
							StoredProc.ParamByName("SCURRENCY").value                = scurrency
							StoredProc.ParamByName("NCREDIT_SUM").value                = 0
							StoredProc.ParamByName("SAGNACC").value                        = agn_strcode
							StoredProc.ParamByName("SSUBDIV").value                        = subdiv
							StoredProc.ParamByName("NDISCOUNT").value                = 0
							StoredProc.ParamByName("NPRICE_TYPE").value                = 0
							StoredProc.ParamByName("NSIGNTAX").value                = 1
							StoredProc.ParamByName("NSAME_NOMN").value                = 0
							StoredProc.ExecProc

							Set ExtraData = nodeNode.selectNodes("ДопРеквизиты/Строки")
							If ExtraData.length > 0 then
								For Each node In ExtraData
									AttribName = node.selectSingleNode("Имя").text
									If AttribName = "Способ закупки" or AttribName = "Количество поданных заявок" or AttribName = "Количество недопущенных заявок" or AttribName = "Начальная цена закупки" or AttribName = "№ закупочной процедуры" then
										Query.SQL.text = "select CODE from docs_props where name = '"&AttribName&"'"
										Query.Open

										StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"
										StoredProc.ParamByName("PROPERTY").value=Query.FieldByname("CODE").value
										StoredProc.ParamByName("UNITCODE").value="Contracts"
										StoredProc.ParamByName("RN_SOTR").value=newRN
										StoredProc.ParamByName("ST_VAL").value=node.selectSingleNode("Значение").text
										StoredProc.ParamByName("NUM_VAL").value=NULL
										StoredProc.ExecProc
									elseif AttribName = "Принадлежность к субъектам мал/ср. бизнеса" then
										If node.selectSingleNode("Значение").text="да" then
											val = 1
										else
											val = 0
										end if
										StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"
										StoredProc.ParamByName("PROPERTY").value="Субъект мал/ср бизн."
										StoredProc.ParamByName("UNITCODE").value="Contracts"
										StoredProc.ParamByName("RN_SOTR").value=newRN
										StoredProc.ParamByName("ST_VAL").value=NULL
										StoredProc.ParamByName("NUM_VAL").value=val
										StoredProc.ExecProc
									end if
								next
							end if
						else
							'ДОБАВИМ ЭТАП К ДОГОВОРУ
							Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&nodeNode.selectSingleNode("БазовыйДоговор").text&"' and docs_prop_rn='104582667' and unitcode='Contracts'"        ' 87456099 - код свойства "Код 1С"
							Query.Open
							UNIT_RN = Query.FieldByname("UNIT_RN").value
							If Query.IsEmpty then
								MyFile.Write(vbTab&"ERROR "&now()&vbTab&" В Парус не найден родительский договор с кодом 1С "&nodeNode.selectSingleNode("БазовыйДоговор").text&" - не могу завести новый этап"&vbNewLine)
							else
								Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&nodeNode.selectSingleNode("Ссылка").text&"' and docs_prop_rn='109199304' and unitcode='ContractsStages'"        ' 109199304 - код свойства "Код 1С"
								Query.Open
								If Query.IsEmpty then
									MyFile.Write(vbTab&"INFO "&now()&vbTab&" В Парус найден родительский договор с кодом 1С "&nodeNode.selectSingleNode("БазовыйДоговор").text&" - завожу новый этап"&vbNewLine)

									'ПОЛУЧИМ ПРЕФИКС НУМЕРАЦИИ ЭТАПА
									Query.SQL.Text="select DOC_PREF, DOC_NUMB from CONTRACTS where RN='"&UNIT_RN&"'"
									Query.Open
									doc_pref = Trim(Query.FieldByname("DOC_PREF").value)
									doc_numb = Trim(Query.FieldByname("DOC_NUMB").value)

									'ПОЛУЧИМ ПОРЯДКОВЫЙ НОМЕР
									StoredProc.StoredProcName="P_STAGES_GETNEXTNUMB"
									StoredProc.ParamByName("NCOMPANY").value                = 42903
									StoredProc.ParamByName("NPRN").value                        = UNIT_RN
									StoredProc.ExecProc
									snumb = StoredProc.ParamByName("SNUMB_MAX").value

									'ДОБАВИМ НОВЫЙ ЭТАП К ДОГОВОРУ
									StoredProc.StoredProcName="P_STAGES_INSERT"
									StoredProc.ParamByName("nCOMPANY").value                = 42903
									StoredProc.ParamByName("nPRN").value                        = UNIT_RN
									StoredProc.ParamByName("SNUMB").value                        = snumb
									StoredProc.ParamByName("NEXT_AGREEMENT").value        = 1
									StoredProc.ParamByName("NSIGN_SUM").value                = 1
									StoredProc.ParamByName("DBEGIN_DATE").value                = ConvDate(nodeNode.selectSingleNode("СрокДействияС").text)
									StoredProc.ParamByName("DEND_DATE").value                = endDate
									StoredProc.ParamByName("SJUR_ACC").value                = jur_strcode
									StoredProc.ParamByName("NSUM_TYPE").value                = sumType                                'расчет суммы
									StoredProc.ParamByName("NSTAGE_SUM").value                = 0
									StoredProc.ParamByName("NSTAGE_SUMTAX").value        = Replace(nodeNode.selectSingleNode("СуммаДоговора").text, ".", ",")
									StoredProc.ParamByName("NSTAGE_SUM_NDS").value        = 0
									StoredProc.ParamByName("NAUTOCALC_SIGN").value        = 1
									StoredProc.ParamByName("SDESCRIPTION").value        = nodeNode.selectSingleNode("ПредметДоговора").text
									StoredProc.ParamByName("SCOMMENTS").value                = nodeNode.selectSingleNode("ПредметДоговора").text
									StoredProc.ParamByName("NFACEACC_EXIST").value        = 0
									StoredProc.ParamByName("SFACEACCCRN").value                = GetFaceAccCat(subdiv)'связать по подразделению
									StoredProc.ParamByName("SAGENT").value                        = agn_abbr
									StoredProc.ParamByName("SFACEACC").value                = doc_pref&"/"&doc_numb&"/"&snumb
									StoredProc.ParamByName("NACC_KIND").value                = acc_kind
									StoredProc.ParamByName("SEXECUTIVE").value                = executive
									StoredProc.ParamByName("SCURRENCY").value                = scurrency
									StoredProc.ParamByName("NCREDIT_SUM").value                = 0
									StoredProc.ParamByName("SAGNACC").value                        = agn_strcode
									StoredProc.ParamByName("SSUBDIV").value                        = subdiv
									StoredProc.ParamByName("NDISCOUNT").value                = 0
									StoredProc.ParamByName("NPRICE_TYPE").value                = 0
									StoredProc.ParamByName("NSIGNTAX").value                = 1
									StoredProc.ParamByName("NSAME_NOMN").value                = 0
									StoredProc.ExecProc
									newRN = StoredProc.ParamByName("nRN").value

									'ЗАПИШЕМ ДАННЫЕ В СВОЙСТВА ДОКУМЕНТА
									StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'ДогКод1С
									StoredProc.ParamByName("PROPERTY").value="ДогЭтапКод1С"
									StoredProc.ParamByName("UNITCODE").value="ContractsStages"
									StoredProc.ParamByName("RN_SOTR").value=newRN
									StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("Ссылка").text)
									StoredProc.ParamByName("NUM_VAL").value=NULL
									StoredProc.ExecProc
								else
									MyFile.Write("INFO "&now()&vbTab&" В Парус найден этап "&comment&" ("&trim(nodeNode.selectSingleNode("Ссылка").text)&") - пропускаю этап."&vbNewLine)
								end if
							end if
						end if
					end if
				else
					MyFile.Write("INFO "&now()&vbTab&" В Парус найден договор "&comment&" ("&trim(nodeNode.selectSingleNode("Ссылка").text)&") - пропускаю договор."&vbNewLine)
					'ЗАПИШЕМ ДАННЫЕ В СВОЙСТВА ДОКУМЕНТА
					StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'ДогКод1С
					StoredProc.ParamByName("PROPERTY").value="ДогКод1С"
					StoredProc.ParamByName("UNITCODE").value="Contracts"
					StoredProc.ParamByName("RN_SOTR").value=old_RN
					StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("Ссылка").text)
					StoredProc.ParamByName("NUM_VAL").value=NULL
					StoredProc.ExecProc
				end if
			end if
		next
		MyFile.Write("INFO "&now()&vbTab&" Договора загружены. Всего обработано: "& object_counter &vbNewLine&vbNewLine)
	else
		MyFile.Write("INFO "&now()&vbTab&" В текущей выгрузке на найдено информации о договорах."&vbNewLine&vbNewLine)
	end if

	'НАХОДИМ УЗЕЛ "СписанияСРасчетногоСчета"'
	Set Outcoming = xmlParser.selectNodes("//СписанияСРасчетногоСчета/Строки")
	If Outcoming.length > 0 then
		object_counter = 0
		For Each nodeNode In Outcoming
			DATE_array                = NULL
			DATE_year                = NULL
			banknumb                = NULL
			DATETIME_array        = NULL
			DATEONLY                = NULL
			docpref                 = NULL
			docnumb                 = NULL
			doctype                 = NULL
			delimiter                = NULL
			docdate                 = NULL
			jur_strcode                = NULL
			agn_abbr                = NULL
			agnaccbik                = NULL
			agnacc                        = NULL
			agn_strcode                = NULL
			typeoper_mnemo        = NULL
			scurrency                = NULL
			newRN                        = NULL
			agn_rn                        = NULL
			SPAY_NOTE                = NULL

			object_counter = object_counter + 1

			Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value like '%"&trim(nodeNode.selectSingleNode("Ссылка").text)&"%' and docs_prop_rn='104582883' and unitcode='BankDocuments'"
			Query.Open
			If Query.IsEmpty then
				MyFile.Write("INFO "&now()&vbTab&" В Парус не найден документ списания ДС с кодом "&trim(nodeNode.selectSingleNode("Ссылка").text)&"  "&nodeNode.selectSingleNode("ВхНомер").text&" "&nodeNode.selectSingleNode("ВхДата").text&" - создаю новый документ."&vbNewLine)

				'ПОЛУЧИМ ГОД ИЗ ДАТЫ
				DATE_array	= Split(nodeNode.selectSingleNode("ВхДата").text, "-")
				DATE_year	= DATE_array(0)

				'ПОЛУЧИМ ПОРЯДКОВЫЙ НОМЕР
				StoredProc.StoredProcName="P_BANKDOCS_GETNEXTNUMB"
				StoredProc.ParamByName("NCOMPANY").value		= 42903
				StoredProc.ParamByName("SJUR_PERS").value		= "НК ТЭЦ"
				StoredProc.ParamByName("DBANK_DOCDATE").value	= ConvDate(nodeNode.selectSingleNode("ВхДата").text)
				StoredProc.ParamByName("SBANK_DOCTYPE").value	= "ИсходП/П"
				StoredProc.ParamByName("SBANK_DOCPREF").value	= DATE_year
				StoredProc.ExecProc
				banknumb = StoredProc.ParamByName("SBANK_NUMB").value

				'ВЫДЕЛИМ ДАТУ ИЗ СТРОКИ ВИДА датаТвремя
				DATETIME_array        = Split(nodeNode.selectSingleNode("Дата").text, "T")
				DATEONLY = DATETIME_array(0)

				'ПОЛУЧИМ ДАННЫЕ ДОГОВОРА, ЕСЛИ В ДОКУМЕНТЕ ТОЛЬКО 1 ДОГОВОР
				Set TableRows = nodeNode.selectNodes("ТабЧасть/Строки")
				If TableRows.length = 1 then
					Set Node = TableRows.nextNode()
					Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&Node.selectSingleNode("Договор").text&"' and docs_prop_rn='104582667' and unitcode='Contracts'"        ' 87456099 - код свойства "Код 1С"
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

					'ПОЛУЧИМ СТАВКУ НДС И ЕЕ ЗНАЧЕНИЕ
					If Node.selectSingleNode("СтавкаНДС").text="БезНДС" or Node.selectSingleNode("СтавкаНДС").text="" then
						TAXrate        = 0
						Tax                = 0
					else
						TAXrate        = Mid(Node.selectSingleNode("СтавкаНДС").text, 4)
						Tax                = Replace(Node.selectSingleNode("СуммаНДС").text, ".", ",")
					End if
				end if

				'НАЙДЕМ МНЕМОКОД НАШЕЙ ОРГАНИЗАЦИИ ПО ЕЕ КОДУ SAP
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - код свойства "Код SAP" (1ИЗМЕНИТЬ ПОСЛЕ ПЕРЕНОСА В ПРОДУКТИВ)
				Query.Open
				Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				SAGENT_FROM = Query.FieldByname("AGNABBR").value
				Query.Close

				'ПОЛУЧИМ НОМЕР СТРОКИ БАНКОВСКОГО СЧЕТА ОРГАНИЗАЦИИ ЧЕРЕЗ НОМЕР ЭТОГО СЧЕТА
				Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & Trim(nodeNode.selectSingleNode("СчетОрганизацииБИК").text) & "'"
				Query.Open
				If not Query.IsEmpty then
					Query.SQL.Text        = "select STRCODE from AGNACC where AGNACC='"& Trim(nodeNode.selectSingleNode("СчетОрганизации").text) &"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='5805775'"
					Query.Open
					jur_strcode                = Query.FieldByname("STRCODE").value
					Query.Close
				else
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Для банковского счета с номером "&nodeNode.selectSingleNode("СчетОрганизации").text&" в Парус не найден БИК "&nodeNode.selectSingleNode("СчетОрганизацииБИК").text&" - делаю останов для отладки."&vbNewLine)
				end if


				'ИЩЕМ КОНТРАГЕНТА И СЧЕТ В ЗАВИСИМОСТИ ОТ РАЗЛИЧНЫХ УСЛОВИЙ
				if nodeNode.selectSingleNode("ВидОперации").text = "ПереводНаДругойСчет" then
					'КОНТРАГЕНТ - ТЭЦ
					Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
					Query.Open
					agn_rn = Query.FieldByname("UNIT_RN").value
					Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value

					'CЧЕТ - ПО ЕГО НОМЕРУ
					agnaccbik                = Trim(nodeNode.selectSingleNode("СчетКонтрагентаБИК").text)
					agnacc                        = Trim(nodeNode.selectSingleNode("СчетКонтрагента").text)
					If not nodeNode.selectSingleNode("СчетКонтрагентаБИК").text="" then
						Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & agnaccbik & "'"
						Query.Open
						bankrn = Query.FieldByname("NRN").value
					else
						bankrn = NULL
					end if
					If not IsNull(bankrn) then
						Query.SQL.Text         = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"&bankrn&"' and AGNRN='"&agn_rn&"'"
						Query.Open
						If not Query.IsEmpty then
							agn_strcode = Query.FieldByname("STRCODE").value
						else
							'НЕ НАШЛИ ПО НОМЕРУ - ПО НАЗВАНИЮ БАНКА
							Query.SQL.Text = "select STRCODE from agnacc where agnrn='5805775' and agnbanks='"&bankrn&"'"
							Query.Open
							agn_strcode = Query.FieldByname("STRCODE").value
						end if
					else
						MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Для банковского счета с номером "&nodeNode.selectSingleNode("СчетКонтрагента").text&" в Парус не найден БИК "&nodeNode.selectSingleNode("СчетКонтрагентаБИК").text&" - делаю останов для отладки."&vbNewLine)
					end if
				else
					'ИЩЕМ КОНТРАГЕНТА
					Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='"&RTrim(nodeNode.selectSingleNode("Контрагент").text)&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
					Query.Open
					agn_rn = Query.FieldByname("UNIT_RN").value
					Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value

					'ИЩЕМ СЧЕТ
					If nodeNode.selectSingleNode("СчетКонтрагента").text = "" then 'ПЕРВЫЙ ПОПАВШИЙСЯ СЧЕТ
						Query.SQL.Text         = "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
						Query.Open
						agn_strcode         = Query.FieldByname("STRCODE").value
					else
						agnaccbik                = Trim(nodeNode.selectSingleNode("СчетКонтрагентаБИК").text)
						agnacc                        = Trim(nodeNode.selectSingleNode("СчетКонтрагента").text)
						If not nodeNode.selectSingleNode("СчетКонтрагентаБИК").text="" then
							Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & agnaccbik & "'"
							Query.Open
							bankrn = Query.FieldByname("NRN").value
						else
							bankrn = NULL
						end if
						If IsNull(bankrn) then
							Query.SQL.Text         = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS is NULL and AGNRN='"&agn_rn&"'"
							Query.Open
							agn_strcode         = Query.FieldByname("STRCODE").value
						else
							Query.SQL.Text         = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"& bankrn &"' and AGNRN='"&agn_rn&"'"
							Query.Open
							agn_strcode         = Query.FieldByname("STRCODE").value
						end if
					end if
				end if

				'НАЙДЕМ ВИД ФИН ОПЕРАЦИИ ПО ЕГО КОДУ 1С
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='р"&nodeNode.selectSingleNode("ВидОперации").text&"' and docs_prop_rn='104582941' and unitcode='TypeOpersPay'"        ' 104582941 - код свойства "Код SAP"
				Query.Open
				Query.SQL.Text="select TYPOPER_MNEMO from DICTOPER where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				typeoper_mnemo = Query.FieldByname("TYPOPER_MNEMO").value
				Query.Close

				'ПОЛУЧИМ НАИМЕНОВАНИЕ ВАЛЮТЫ ПО ЕЕ КОДУ
				Query.SQL.Text = "select INTCODE from curnames where curcode='"&nodeNode.selectSingleNode("ВалютаДокумента").text&"'"
				Query.Open
				If Query.IsEmpty then
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Валюта с кодом "&nodeNode.selectSingleNode("ВалютаДокумента").text&" из договора <<"&comment&">> не найдена. Используется значение по-умолчанию - RUR")
					scurrency = "RUR"
				else
					scurrency = nodeNode.selectSingleNode("ВалютаДокумента").text
				end if
				Query.Close

				'СОЗДАЕМ ЗАПИСЬ О НОВОМ ДОКУМЕНТЕ СПИСАНИЯ ДС
				StoredProc.StoredProcName="P_BANKDOCSACC_INSERT"
				StoredProc.ParamByName("nCOMPANY").value                = 42903                        'код подразделения
				StoredProc.ParamByName("nCRN").value                        = 104583621                'код каталога 1ИЗМЕНИТЬ ПОСЛЕ ПЕРЕНОСА В ПРОДУКТИВ
				StoredProc.ParamByName("SBANK_TYPEDOC").value        = "ИсходП/П"
				StoredProc.ParamByName("SBANK_PREFDOC").value        = DATE_year
				StoredProc.ParamByName("SBANK_NUMBDOC").value        = banknumb
				StoredProc.ParamByName("DBANK_DATEDOC").value        = ConvDate(DATEONLY)
				StoredProc.ParamByName("SVALID_TYPEDOC").value        = doctype
				StoredProc.ParamByName("SVALID_NUMBDOC").value        = docpref & delimiter & docnumb
				StoredProc.ParamByName("DVALID_DATEDOC").value        = docdate
				StoredProc.ParamByName("SFROM_NUMB").value                = nodeNode.selectSingleNode("ВхНомер").text
				StoredProc.ParamByName("DFROM_DATE").value                = ConvDate(nodeNode.selectSingleNode("ВхДата").text)
				StoredProc.ParamByName("SAGENT_FROM").value                = SAGENT_FROM
				StoredProc.ParamByName("SAGENTF_ACC").value                = jur_strcode
				StoredProc.ParamByName("SAGENT_TO").value                = agn_abbr
				StoredProc.ParamByName("SAGENTT_ACC").value                = agn_strcode
				StoredProc.ParamByName("STYPE_OPER").value                = typeoper_mnemo
				StoredProc.ParamByName("SPAY_INFO").value                = nodeNode.selectSingleNode("НазначениеПлатежа").text
				StoredProc.ParamByName("SPAY_NOTE").value                = SPAY_NOTE
				StoredProc.ParamByName("NPAY_SUM").value                = Replace(nodeNode.selectSingleNode("СуммаДокумента").text, ".", ",")
				StoredProc.ParamByName("NTAX_SUM").value                = Tax
				StoredProc.ParamByName("NPERCENT_TAX_SUM").value= TAXrate
				StoredProc.ParamByName("SCURRENCY").value                = scurrency
				StoredProc.ParamByName("SJUR_PERS").value                = "НК ТЭЦ"
				StoredProc.ParamByName("NUNALLOTTED_SUM").value        = 0
				StoredProc.ParamByName("NIS_ADVANCE").value                = 0
				StoredProc.ExecProc
				newRN = StoredProc.ParamByName("nRN").value

				'ЗАПИШЕМ ДАННЫЕ В СВОЙСТВА ДОКУМЕНТА
				StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'ДДСКод1С
				StoredProc.ParamByName("PROPERTY").value="ДДСКод1С"
				StoredProc.ParamByName("UNITCODE").value="BankDocuments"
				StoredProc.ParamByName("RN_SOTR").value=newRN
				StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("Ссылка").text)
				StoredProc.ParamByName("NUM_VAL").value=NULL
				StoredProc.ExecProc
			else
				MyFile.Write("INFO "&now()&vbTab&" В Парус найден документ списания ДС с кодом "&trim(nodeNode.selectSingleNode("Ссылка").text)&"  "&nodeNode.selectSingleNode("ВхНомер").text&" "&nodeNode.selectSingleNode("ВхДата").text&" - пропускаю документ."&vbNewLine)
			End if
			Query.Close
		next
		MyFile.Write("INFO "&now()&vbTab&" Документы списания с расчетных счетов загружены. Всего обработано: "& object_counter & vbNewLine&vbNewLine)
	end if

	'НАХОДИМ УЗЕЛ "ПоступленияНаРасчетныйСчет"'
	Set Incoming = xmlParser.selectNodes("//ПоступленияНаРасчетныйСчет/Строки")
	If Incoming.length > 0 then
		object_counter = 0
		For Each nodeNode In Incoming
			DATE_array                = NULL
			DATE_year                = NULL
			banknumb                = NULL
			DATETIME_array        = NULL
			DATEONLY                = NULL
			docpref                 = NULL
			docnumb                 = NULL
			doctype                 = NULL
			docdate                 = NULL
			jur_strcode                = NULL
			agn_abbr                = NULL
			agnaccbik                = NULL
			agnacc                        = NULL
			agn_strcode                = NULL
			typeoper_mnemo        = NULL
			scurrency                = NULL
			newRN                        = NULL
			agn_rn                        = NULL
			SPAY_NOTE                = NULL
			bankrn                        = NULL

			object_counter = object_counter + 1

			Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value like '%"&trim(nodeNode.selectSingleNode("Ссылка").text)&"%' and docs_prop_rn='104582883' and unitcode='BankDocuments'"
			Query.Open
			If Query.IsEmpty then
				MyFile.Write("INFO "&now()&vbTab&" В Парус не найден документ поступления ДС с кодом "&trim(nodeNode.selectSingleNode("Ссылка").text)&" "&nodeNode.selectSingleNode("ВхНомер").text&" "&nodeNode.selectSingleNode("ВхДата").text&" - создаю новый документ."&vbNewLine)

				'ПОЛУЧИМ ГОД ИЗ ДАТЫ
				DATE_array        = Split(nodeNode.selectSingleNode("ВхДата").text, "-")
				DATE_year        = DATE_array(0)

				'ПОЛУЧИМ ПОРЯДКОВЫЙ НОМЕР
				StoredProc.StoredProcName="P_BANKDOCS_GETNEXTNUMB"
				StoredProc.ParamByName("NCOMPANY").value                = 42903
				StoredProc.ParamByName("SJUR_PERS").value                = "НК ТЭЦ"
				StoredProc.ParamByName("DBANK_DOCDATE").value        = ConvDate(nodeNode.selectSingleNode("ВхДата").text)
				StoredProc.ParamByName("SBANK_DOCTYPE").value        = "ВходП/П"
				StoredProc.ParamByName("SBANK_DOCPREF").value        = DATE_year
				StoredProc.ExecProc
				banknumb = StoredProc.ParamByName("SBANK_NUMB").value

				'ВЫДЕЛИМ ДАТУ ИЗ СТРОКИ ВИДА датаТвремя
				DATETIME_array        = Split(nodeNode.selectSingleNode("Дата").text, "T")
				DATEONLY = DATETIME_array(0)

				'ПОЛУЧИМ ДАННЫЕ ДОГОВОРА, ЕСЛИ В ДОКУМЕНТЕ ТОЛЬКО 1 ДОГОВОР
				docpref = NULL
				docnumb = NULL
				doctype = NULL
				docdate = NULL
				delimiter = NULL
				TAXrate        = 0
				Tax                = 0
				Set TableRows = nodeNode.selectNodes("ТабЧасть/Строки")
				If TableRows.length = 1 then
					Set Node = TableRows.nextNode()
					Query.SQL.Text="select UNIT_RN from docs_props_vals where str_value='"&Node.selectSingleNode("Договор").text&"' and docs_prop_rn='87456099' and unitcode='Contracts'"        ' 87456099 - код свойства "ДогКод1С"
					Query.Open
					If not Query.IsEmpty then
						Query.SQL.Text="select DOC_TYPE, DOC_PREF, DOC_NUMB, DOC_DATE from CONTRACTS where RN='"&Query.FieldByname("UNIT_RN").value&"'"
						Query.Open
						If docdate="01-янв-0001" then 'дата 01-янв-0001 в банковских документах не дает сформировать платежи по счетам
							docdate = ""
						else
							docdate = Query.FieldByname("DOC_DATE").value
						end if
						docpref = LTrim(Query.FieldByname("DOC_PREF").value)
						docnumb = LTrim(Query.FieldByname("DOC_NUMB").value)
						delimiter = "-"
						Query.SQL.Text="select DOCCODE from DOCTYPES where RN='"&Query.FieldByname("DOC_TYPE").value&"'"
						Query.Open
						doctype = Query.FieldByname("DOCCODE").value
						Query.Close
					end if

					'ПОЛУЧИМ СТАВКУ НДС И ЕЕ ЗНАЧЕНИЕ
					If Node.selectSingleNode("СтавкаНДС").text="БезНДС" or Node.selectSingleNode("СтавкаНДС").text="" then
						TAXrate        = 0
						Tax                = 0
					else
						TAXrate        = Mid(Node.selectSingleNode("СтавкаНДС").text, 4)
						Tax                = Replace(Node.selectSingleNode("СуммаНДС").text, ".", ",")
					End if
				end if

				'НАЙДЕМ МНЕМОКОД НАШЕЙ ОРГАНИЗАЦИИ ПО ЕЕ КОДУ SAP
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"        ' 105510718 - код свойства "Код SAP" (1ИЗМЕНИТЬ ПОСЛЕ ПЕРЕНОСА В ПРОДУКТИВ)
				Query.Open
				Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				SAGENT_TO = Query.FieldByname("AGNABBR").value
				Query.Close

				'ПОЛУЧИМ НОМЕР СТРОКИ БАНКОВСКОГО СЧЕТА ОРГАНИЗАЦИИ ЧЕРЕЗ НОМЕР ЭТОГО СЧЕТА
				Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & Trim(nodeNode.selectSingleNode("СчетОрганизацииБИК").text) & "'"
				Query.Open
				If not Query.IsEmpty then
					Query.SQL.Text        = "select STRCODE from AGNACC where AGNACC='"& Trim(nodeNode.selectSingleNode("СчетОрганизации").text) &"' and AGNBANKS='" & Query.FieldByname("NRN").value & "' and AGNRN='5805775'"
					Query.Open
					jur_strcode                = Query.FieldByname("STRCODE").value
					Query.Close
				else
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Для банковского счета с номером "&nodeNode.selectSingleNode("СчетОрганизации").text&" в Парус не найден БИК "&nodeNode.selectSingleNode("СчетОрганизацииБИК").text&" - делаю останов для отладки."&vbNewLine)
				end if

				'ИЩЕМ КОНТРАГЕНТА И СЧЕТ В ЗАВИСИМОСТИ ОТ РАЗЛИЧНЫХ УСЛОВИЙ
				if nodeNode.selectSingleNode("ВидОперации").text = "ПереводСДругогоСчета" then
					'КОНТРАГЕНТ - ТЭЦ
					Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='0000112063' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
					Query.Open
					agn_rn = Query.FieldByname("UNIT_RN").value
					Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
					Query.Open
					agn_abbr = Query.FieldByname("AGNABBR").value

					if nodeNode.selectSingleNode("Контрагент").text="" or nodeNode.selectSingleNode("СчетКонтрагента").text = "" then
						SPAY_NOTE = "Контрагент (код SAP) "&nodeNode.selectSingleNode("Контрагент").text&", счет "&nodeNode.selectSingleNode("СчетКонтрагента").text
						agn_abbr = NULL
						agn_strcode = NULL
					else
						'CЧЕТ - ПО ЕГО НОМЕРУ
						agnaccbik                = Trim(nodeNode.selectSingleNode("СчетКонтрагентаБИК").text)
						agnacc                        = Trim(nodeNode.selectSingleNode("СчетКонтрагента").text)
						If not nodeNode.selectSingleNode("СчетКонтрагентаБИК").text="" then
							Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & agnaccbik & "'"
							Query.Open
							bankrn = Query.FieldByname("NRN").value
						else
							bankrn = NULL
						end if

						If not IsNull(bankrn) then
							Query.SQL.Text         = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"&bankrn&"' and AGNRN='"&agn_rn&"'"
							Query.Open
							If not Query.IsEmpty then
								agn_strcode = Query.FieldByname("STRCODE").value
							else
								'НЕ НАШЛИ ПО НОМЕРУ - ПО НАЗВАНИЮ БАНКА
								Query.SQL.Text = "select STRCODE from agnacc where agnrn='5805775' and agnbanks='"&bankrn&"'"
								Query.Open
								agn_strcode = Query.FieldByname("STRCODE").value
							end if
						else
							MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Для банковского счета с номером "&nodeNode.selectSingleNode("СчетКонтрагента").text&" в Парус не найден БИК "&nodeNode.selectSingleNode("СчетКонтрагентаБИК").text&" - делаю останов для отладки."&vbNewLine)
						end if
					end if
				else
					if nodeNode.selectSingleNode("Контрагент").text = "" then
						SPAY_NOTE = "Контрагент (код SAP) "&nodeNode.selectSingleNode("Контрагент").text&", счет "&nodeNode.selectSingleNode("СчетКонтрагента").text
						agn_abbr        = NULL
						agn_strcode = NULL
					else
						'ИЩЕМ КОНТРАГЕНТА
						Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='"&nodeNode.selectSingleNode("Контрагент").text&"' and docs_prop_rn='105510718' and unitcode='AGNLIST'"
						Query.Open
						agn_rn = Query.FieldByname("UNIT_RN").value
						Query.SQL.Text="select AGNABBR from AGNLIST where RN='"&agn_rn&"'"
						Query.Open
						agn_abbr = Query.FieldByname("AGNABBR").value

						'ИЩЕМ СЧЕТ
						If nodeNode.selectSingleNode("СчетКонтрагента").text = "" then        'ПЕРВЫЙ ПОПАВШИЙСЯ СЧЕТ
							Query.SQL.Text         = "select STRCODE from AGNACC where AGNRN='"&agn_rn&"'"
							Query.Open
							agn_strcode = Query.FieldByname("STRCODE").value
						else
							agnaccbik                = Trim(nodeNode.selectSingleNode("СчетКонтрагентаБИК").text)
							agnacc                        = Trim(nodeNode.selectSingleNode("СчетКонтрагента").text)
							If not nodeNode.selectSingleNode("СчетКонтрагентаБИК").text="" then
								Query.SQL.Text        = "select NRN from v_agnbanks where SCODE='БАНК_" & agnaccbik & "'"
								Query.Open
								bankrn = Query.FieldByname("NRN").value
							else
								bankrn = NULL
							end if
							If IsNull(bankrn) then
								Query.SQL.Text         = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS is NULL and AGNRN='"&agn_rn&"'"
								Query.Open
								agn_strcode         = Query.FieldByname("STRCODE").value
							else
								Query.SQL.Text         = "select STRCODE from AGNACC where AGNACC='"&agnacc&"' and AGNBANKS='"& bankrn &"' and AGNRN='"&agn_rn&"'"
								Query.Open
								agn_strcode         = Query.FieldByname("STRCODE").value
							end if
						end if
					end if
				end if

				'НАЙДЕМ ВИД ФИН ОПЕРАЦИИ ПО ЕГО КОДУ 1С
				Query.Sql.Text="select UNIT_RN from docs_props_vals where str_value='п"&nodeNode.selectSingleNode("ВидОперации").text&"' and docs_prop_rn='104582941' and unitcode='TypeOpersPay'"        ' 104582941 - код свойства "Код SAP"
				Query.Open
				Query.SQL.Text="select TYPOPER_MNEMO from DICTOPER where RN='"&Query.FieldByname("UNIT_RN").value&"'"
				Query.Open
				typeoper_mnemo = Query.FieldByname("TYPOPER_MNEMO").value
				Query.Close

				'ПОЛУЧИМ НАИМЕНОВАНИЕ ВАЛЮТЫ ПО ЕЕ КОДУ
				Query.SQL.Text = "select INTCODE from curnames where curcode='"&nodeNode.selectSingleNode("ВалютаДокумента").text&"'"
				Query.Open
				If Query.IsEmpty then
					MyFile.Write(vbTab&"ERROR "&now()&vbTab&" Валюта с кодом "&nodeNode.selectSingleNode("ВалютаДокумента").text&" из договора <<"&comment&">> не найдена. Используется значение по-умолчанию - RUR")
					scurrency = "RUR"
				else
					scurrency = nodeNode.selectSingleNode("ВалютаДокумента").text
				end if
				Query.Close

				'СОЗДАЕМ ЗАПИСЬ О НОВОМ ДОКУМЕНТЕ ПОСТУПЛЕНИЯ ДС
				StoredProc.StoredProcName="P_BANKDOCSACC_INSERT"
				StoredProc.ParamByName("nCOMPANY").value                = 42903                        'код подразделения
				StoredProc.ParamByName("nCRN").value                        = 104583621                'код каталога 1ИЗМЕНИТЬ ПОСЛЕ ПЕРЕНОСА В ПРОДУКТИВ
				StoredProc.ParamByName("SBANK_TYPEDOC").value        = "ВходП/П"
				StoredProc.ParamByName("SBANK_PREFDOC").value        = DATE_year
				StoredProc.ParamByName("SBANK_NUMBDOC").value        = banknumb
				StoredProc.ParamByName("DBANK_DATEDOC").value        = ConvDate(DATEONLY)
				StoredProc.ParamByName("SVALID_TYPEDOC").value        = doctype
				StoredProc.ParamByName("SVALID_NUMBDOC").value        = docpref & delimiter & docnumb
				StoredProc.ParamByName("DVALID_DATEDOC").value        = docdate
				StoredProc.ParamByName("SFROM_NUMB").value                = nodeNode.selectSingleNode("ВхНомер").text
				StoredProc.ParamByName("DFROM_DATE").value                = ConvDate(nodeNode.selectSingleNode("ВхДата").text)
				StoredProc.ParamByName("SAGENT_FROM").value                = agn_abbr
				StoredProc.ParamByName("SAGENTF_ACC").value                = agn_strcode
				StoredProc.ParamByName("SAGENT_TO").value                = SAGENT_TO
				StoredProc.ParamByName("SAGENTT_ACC").value                = jur_strcode
				StoredProc.ParamByName("STYPE_OPER").value                = typeoper_mnemo
				StoredProc.ParamByName("SPAY_INFO").value                = nodeNode.selectSingleNode("НазначениеПлатежа").text
				StoredProc.ParamByName("SPAY_NOTE").value                = SPAY_NOTE
				StoredProc.ParamByName("NPAY_SUM").value                = Replace(nodeNode.selectSingleNode("СуммаДокумента").text, ".", ",")
				StoredProc.ParamByName("NTAX_SUM").value                = Tax
				StoredProc.ParamByName("NPERCENT_TAX_SUM").value= TAXrate
				StoredProc.ParamByName("SCURRENCY").value                = scurrency
				StoredProc.ParamByName("SJUR_PERS").value                = "НК ТЭЦ"
				StoredProc.ParamByName("NUNALLOTTED_SUM").value        = 0
				StoredProc.ParamByName("NIS_ADVANCE").value                = 0
				StoredProc.ExecProc
				newRN = StoredProc.ParamByName("nRN").value

				'ЗАПИШЕМ ДАННЫЕ В СВОЙСТВА ДОКУМЕНТА
				StoredProc.StoredProcName="P_KOD_KONTR_1C_TO_PARUS"         'ДДСКод1С
				StoredProc.ParamByName("PROPERTY").value="ДДСКод1С"
				StoredProc.ParamByName("UNITCODE").value="BankDocuments"
				StoredProc.ParamByName("RN_SOTR").value=newRN
				StoredProc.ParamByName("ST_VAL").value=trim(nodeNode.selectSingleNode("Ссылка").text)
				StoredProc.ParamByName("NUM_VAL").value=NULL
				StoredProc.ExecProc
			else
				MyFile.Write("INFO "&now()&vbTab&" В Парус найден документ списания ДС с кодом "&trim(nodeNode.selectSingleNode("Ссылка").text)&"  "&nodeNode.selectSingleNode("ВхНомер").text&" "&nodeNode.selectSingleNode("ВхДата").text&" - пропускаю документ."&vbNewLine)
			End if
			Query.Close
		next
		MyFile.Write("INFO "&now()&vbTab&" Документы поступления на расчетные счета загружены.  Всего обработано: "& object_counter & vbNewLine)
	end if

	'ГОТОВИМ ФАЙЛ ОТВЕТА
	Set oldReply = CreateObject("Msxml2.DOMDocument")
	oldReply.async = False
	oldReply.load "\\10.130.32.52\Tatneft\Mess_20100_UH.xml"
	' Проверяем на ошибки загрузки
	If oldReply.parseError.errorCode Then
		MsgBox oldReply.parseError.Reason
	End If
	If not oldReply.parseError.errorCode = -2146697210 then
		oldReplyNumber = cInt(oldReply.selectSingleNode("/Данные/НомерОтправленногоСообщения").text)+1
	else
		oldReplyNumber = 1
	End if
	Set newReply = CreateObject("Msxml2.DOMDocument")
	newReply.appendChild(newReply.createProcessingInstruction("xml", "version='1.0' encoding='utf-8'"))
	Set rootNode = newReply.appendChild( newReply.createElement("Данные") )
	rootNode.setAttribute "xmlns", "http://localhost/ExchangeUH_FileResponse"
	rootNode.setAttribute "xmlns:xs", "http://www.w3.org/2001/XMLSchema"
	rootNode.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
	Set subNode = rootNode.appendChild(newReply.createElement("НомерОтправленногоСообщения"))
	subNode.text = oldReplyNumber
	Set subNode = rootNode.appendChild(newReply.createElement("НомерПринятогоСообщения"))
	subNode.text = xmlParser.selectSingleNode("//НомерОтправленногоСообщения").text
	newReply.save("\\10.130.32.52\Tatneft\Mess_20100_UH.xml")
	'newReply.save("\\10.130.32.52\Tatneft\Exch_logs\Mess_20100_UH.xml")

	'ИМПОРТ ЗАВЕРШЕН
	MyFile.Write(vbTab&now()&vbTab&" Импорт данных из 1С:УХ в ИСУ Парус успешно завершен."&vbNewLine)
	MsgBox "Импорт данных завершен"&vbNewLine&"Подробности в журнале 1C-Parus_exchange.log"
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
	DATE_array        = Split(DateToConvert, "-")
	DATE_year        = DATE_array(0)
	DATE_month        = DATE_array(1)
	DATE_day        = DATE_array(2)
	ConvDate        = DATE_day&"."&DATE_month&"."&DATE_year
End Function

Function GetFaceAccCat(subdiv)
	Select Case subdiv
		Case "НкТЭЦ.04"
		CatRN = 12446958
		
		Case "НкТЭЦ.08"
		CatRN = 4470760
		
		Case "НкТЭЦ.11"
		CatRN = 52315580
		
		Case "НкТЭЦ.13.02"
		CatRN = 4470368
		
		Case "НкТЭЦ.13.02.01"
		CatRN = 4470368
		
		Case "НкТЭЦ.13.02.02"
		CatRN = 4470368
		
		Case "НкТЭЦ.13.04"
		CatRN = 4471005
		
		Case "НкТЭЦ.13.05"
		CatRN = 4471054
		
		Case "НкТЭЦ.13.05.01"
		CatRN = 4471054
		
		Case "НкТЭЦ.13.06"
		CatRN = 4470907
		
		Case "НкТЭЦ.13.09"
		CatRN = 4470417
		
		Case "НкТЭЦ.13.10"
		CatRN = 4470711
		
		Case "НкТЭЦ.13.11"
		CatRN = 4473896
		
		Case "НкТЭЦ.13.12.01"
		CatRN = 4470809
		
		Case "НкТЭЦ.13.15"
		CatRN = 4470809
		
		Case "НкТЭЦ.13.16"
		CatRN = 119088138
		
		Case "НкТЭЦ.13.18"
		CatRN = 4470564
		
		Case "НкТЭЦ.15"
		CatRN = 4472720
		
		Case "НкТЭЦ.17"
		CatRN = 12436655
		
		Case "НкТЭЦ.19"
		CatRN = 111665856
		
		Case "НкТЭЦ.20"
		CatRN = 4473896
		
		Case "НкТЭЦ.21"
		CatRN = 17555197
		
		Case "НкТЭЦ.22"
		CatRN = 4472720
		
		Case "НкТЭЦ.23"
		CatRN = 4472622
		
		Case "НкТЭЦ.24"
		CatRN = 4470319
		
		Case else
		CatRN = 44789479
	End Select
	Query.SQL.Text="select NAME from ACATALOG where RN='"&CatRN&"'"
	Query.Open
	if not Query.IsEmpty then
		GetFaceAccCat = Query.FieldByname("NAME").value
	else
		GetFaceAccCat = "test"
	end if
	
End Function

Function GetContractSubdivPref(subdiv)
	select case subdiv
	
		Case "НкТЭЦ.01"
		doc_pref2 = "374"
	
		Case "НкТЭЦ.02"
		doc_pref2 = "301"
		
		Case "НкТЭЦ.03"
		doc_pref2 = "314"
		
		Case "НкТЭЦ.04"
		doc_pref2 = "379"
		
		Case "НкТЭЦ.05"
		doc_pref2 = "328"
		
		Case "НкТЭЦ.06"
		doc_pref2 = "321"
				
		Case "НкТЭЦ.07"
		doc_pref2 = "361"
		
		Case "НкТЭЦ.08"
		doc_pref2 = "147"
		
		Case "НкТЭЦ.11"
		doc_pref2 = "398"
		
		Case "НкТЭЦ.13.02"
		doc_pref2 = "102"
		
		Case "НкТЭЦ.13.03"
		doc_pref2 = "101"
		
		Case "НкТЭЦ.13.04"
		doc_pref2 = "248"
		
		Case "НкТЭЦ.13.05"
		doc_pref2 = "107"
		
		Case "НкТЭЦ.13.06"
		doc_pref2 = "110"
		
		Case "НкТЭЦ.13.07"
		doc_pref2 = "112"
		
		Case "НкТЭЦ.13.08"
		doc_pref2 = "010"
		
		Case "НкТЭЦ.13.09"
		doc_pref2 = "237"
		
		Case "НкТЭЦ.13.10"
		doc_pref2 = "119"
		
		Case "НкТЭЦ.13.11"
		doc_pref2 = "118"
		
		Case "НкТЭЦ.13.12.02"
		doc_pref2 = "106"
		
		Case "НкТЭЦ.13.13"
		doc_pref2 = "251"
		
		Case "НкТЭЦ.13.15"
		doc_pref2 = "115"
		
		Case "НкТЭЦ.13.16"
		doc_pref2 = "288"
		
		Case "НкТЭЦ.13.18"
		doc_pref2 = "109"
		
		Case "НкТЭЦ.14"
		doc_pref2 = "328м"
		
		Case "НкТЭЦ.15"
		doc_pref2 = "381"
		
		Case "НкТЭЦ.19"
		doc_pref2 = "103"
		
		Case "НкТЭЦ.20"
		doc_pref2 = "023"
		
		Case "НкТЭЦ.21"
		doc_pref2 = "115м"
		
		Case "НкТЭЦ.22"
		doc_pref2 = "090"		
						
		Case "НкТЭЦ.23"
		doc_pref2 = "161"
		
		Case "НкТЭЦ.24"
		doc_pref2 = "091"
		
		Case "Профком"
		doc_pref2 = "707"
	end select
	GetContractSubdivPref = doc_pref2
end function