$include pathxls
        b = open("buffer")
        getfirst(b)
        date_start = trim('��ப�_2'.b)
        date_finish = trim('��ப�_3'.b)
        flag_two = 0
        if substr(date_start,1,4) != substr(date_finish,1,4) then
          flag_two = 1
        end if

        year = substr('��ப�_3'.b,1,4)
        month = substr('��ப�_3'.b,5,2)
        year_month = substr('��ப�_3'.b,1,6)

        name_months = keyarray
        name_months["01"] = array(3)
        name_months["01"][1] = "������"
        name_months["01"][2] = "ﭢ��"
        name_months["01"][3] = 31

        name_months["02"] = array(3)
        name_months["02"][1] = "���ࠫ�"
        name_months["02"][2] = "䥢ࠫ�"
        name_months["02"][3] = 28

        name_months["03"] = array(3)
        name_months["03"][1] = "����"
        name_months["03"][2] = "����"
        name_months["03"][3] = 31

        name_months["04"] = array(3)
        name_months["04"][1] = "��५�"
        name_months["04"][2] = "��५�"
        name_months["04"][3] = 30

        name_months["05"] = array(3)
        name_months["05"][1] = "���"
        name_months["05"][2] = "���"
        name_months["05"][3] = 31

        name_months["06"] = array(3)
        name_months["06"][1] = "���"
        name_months["06"][2] = "�"
        name_months["06"][3] = 30

        name_months["07"] = array(3)
        name_months["07"][1] = "���"
        name_months["07"][2] = "�"
        name_months["07"][3] = 31

        name_months["08"] = array(3)
        name_months["08"][1] = "������"
        name_months["08"][2] = "������"
        name_months["08"][3] = 31

        name_months["09"] = array(3)
        name_months["09"][1] = "�������"
        name_months["09"][2] = "ᥭ���"
        name_months["09"][3] = 30

        name_months["10"] = array(3)
        name_months["10"][1] = "������"
        name_months["10"][2] = "�����"
        name_months["10"][3] = 31

        name_months["11"] = array(3)
        name_months["11"][1] = "�����"
        name_months["11"][2] = "����"
        name_months["11"][3] = 30

        name_months["12"] = array(3)
        name_months["12"][1] = "�������"
        name_months["12"][2] = "������"
        name_months["12"][3] = 31

        people = keyarray

        org = open("organiz",-2)
        getfirst(org)
        tr_plan = 0
        o_plan = 0

        mon = open("months")

        if find(mon,2,2,num(year),num(month)) = 0
          if substr(date_start,5,2) = "01"
            date_b = "26.12." + str(num(substr(date_start,1,4)) - 1,1,0)
            date_btrv = "18.12." + str(num(substr(date_start,1,4)) - 1,1,0)
          else
            if num(substr(date_start,5,2)) - 1 < 10
              date_b = "26.0" + str(num(substr(date_start,5,2)) - 1,1,0) + "." + substr(date_start,1,4)
              date_btrv = "18.0" + str(num(substr(date_start,5,2)) - 1,1,0) + "." + substr(date_start,1,4)
            else
              date_b = "26." + str(num(substr(date_start,5,2)) - 1,1,0) + "." + substr(date_start,1,4)
              date_btrv = "18." + str(num(substr(date_start,5,2)) - 1,1,0) + "." + substr(date_start,1,4)
            end if
          end if
          date_e = "25." + substr(date_finish,5,2) + "." + substr(date_finish,1,4)
          print date_b," // ",date_e
        else
          w_message("�訡��!","���������� ᮡ��� ����� �� 㪠����� �����, ⠪ ��� �� �� ��⨢�஢�� � ����ன���!",0)
          end
        end if

        w_open(300,250,270,160,"������� �롮�")
        w_button(6,80,15,100,25,"����")
        w_button(7,80,50,100,25,"����")
        w_button(10,80,85,100,25,"�⬥��")

        po = open("plan_otchet")
        svod_plan = keyarray

function plan_fact
        result = 1
        local i
        a = regex_split("[:;]",'���祭��'.po)
        for i = 1 to a[0]
          if strpos("R",a[i]) != 0 | strpos("C",a[i]) != 0
            if keyfind(svod_plan,a[i]) = 0
              svod_plan[a[i]] = a[i + 1]
              i = i + 1
            else
              w_message("�訡��!","��������� ���� � �祥�! " + a[i],0)
            end if
          end if
        next
end function

        r = w_show
        if r = 6
          w_close
          if size(po) = 0
            id_po = 1
          else
            getlast(po)
            id_po = 'ID'.po + 1
          end if
          if find(po,2,2,year,month) = 0
            button = w_message("��������!","�����㦥��, �� ࠭�� �ନ஢���� ���� �� ����� �����! �������� ����� �� ����� ���묨?",4)
            plan_fact
            flag_pos = 1
            if button = 6
              flag_insert = 1
            else
              flag_insert = 0
            end if
          else
            flag_insert = 1
            flag_pos = 0
          end if
        else if r = 7
          w_close
          if find(po,2,2,year,month) = 0
            flag_pos = 1
            plan_fact
          else
            flag_pos = 0
            w_message("��������!","�����㦥��, �� ���� �� �ନ஢���� �� ������� ������⥫�! ���� �ᯮ�짮���� � ����� 䠪��᪨� �����!",0)
          end if
        else
          w_close
          w_message("��������!","�믮������ ���� �⬥����!",0)
          end
        end if


function kol_chasov(t1,t2)
        local hour_s,hour_f,min_s,min_f,sec_s,sec_f,i,kol

        hour_s = ""
        hour_f = ""
        min_s = ""
        min_f = ""
        sec_s = ""
        sec_f = ""
        a = regex_split("[:]",t1)
        b = regex_split("[:]",t2)
        kol = 0
        for i = 1 to a[0]
          if strlen(trim(a[i])) != 0 then
            kol = kol + 1
            if kol = 1
              hour_s = a[i]
            else if kol = 2
              min_s = a[i]
            else if kol = 3
              sec_s = a[i]
            end if
          end if
        next

        kol = 0
        for i = 1 to b[0]
          if strlen(trim(b[i])) != 0 then
            kol = kol + 1
            if kol = 1
              hour_f = b[i]
            else if kol = 2
              min_f = b[i]
            else if kol = 3
              sec_f = b[i]
            end if
          end if
        next

        if strlen(trim(hour_s)) != 0 & strlen(trim(hour_f)) != 0 & strlen(trim(min_s)) != 0 & strlen(trim(min_f)) != 0
          result = num(hour_f) - num(hour_s)
          result = result + ( num(min_f) - num(min_s) ) / 60
        else
          result = 0
        end if
end function

function fio_ini(string)
          local fio,sch,i
          fio = ""
          a = regex_split("[ \.]",trim(string))
          sch = 0
          for i = 1 to a[0]
            if strlen(trim(a[i])) != 0 then
              sch = sch + 1
              if sch = 1
                fio = fio + trim(a[i]) + " "
              else if sch = 2 | sch = 3
                fio = fio + substr(trim(a[i]),1,1) + "."
              end if
            end if
          next
          result = fio
end function

        vid = open("vid_oborud",-2)
        itogi = array(30)
        for i = 1 to 30
          itogi[i] = 0
        next

        stops = keyarray

        dde_open("EXCEL")
        dde_execute("[NEW()]")
        dde_execute("[OPEN(" + chr(34) + pathxls + "����-����.xls" + chr(34) + ")]")
        dde_execute("[APP.MINIMIZE()]")

        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from OST_OBORUD where ""���_�����"" IS NOT NULL AND (""��_���_���"" IS NULL OR ""��_���_���""='0')" )
        kolvo = sz["Kolvo"]
        close(sz)
rem     ���� ��

        stanki_to = keyarray

        if kolvo > 0
          ns = sql_query(Upper(getpath),"SELECT * from OST_OBORUD where ""���_�����"" IS NOT NULL AND (""��_���_���"" IS NULL OR ""��_���_���""='0') ORDER BY ""���_��ꥪ�"" ASC, ""��⨪�"" ASC")
          dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "��" + chr(34) + ")]")
          dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
          dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",9,FALSE,FALSE,FALSE,FALSE,1)]")
          dde_execute("[CLEAR()]")
          dde_execute("[BORDER(0,0,0,0,0)]")
          dde_poke("R2C1","�⤥�� �28, ᮣ��᭮ �������� ��䨪� �� � " + name_months[month][2] + " ����� " + year + " �.")

          count_wp = 0

          w_progress("���� ��...",0,kolvo)
          getfirst(ns)
          while( eof(ns) = 0 )
            count_wp = count_wp + 1
            w_progress(count_wp)
            id_oborud = ns["ID_�����"]
rem            print "new stanki ",ns["���_�����"]
            rm = sql_query(Upper(getpath),"SELECT * from WORK_OBORUD where ""ID_�����""='" + str(id_oborud,1,0) + "' AND SUBSTRING(""���_�_��"" FROM 1 FOR 4)='" + year + "' ORDER BY ""ID_ࠡ_��"" ASC")
            getfirst(rm)
            while( eof(rm) = 0 )
              if substr(rm["���_�_��"],1,6) = year_month then
rem                print rm["���_�_��"],rm["ID_�����"],rm["ID_ࠡ_��"],rm["ID_ࠡ_�"]
                if strlen(trim(rm["�⠯_��"])) != 0 then
                  if strlen(trim(ns["���_��ꥪ�"])) != 0
                    if keyfind(stanki_to,ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]) = 0
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]] = array(20)
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][1] = ns["���_��ꥪ�"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][2] = ns["���_�����"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][3] = ns["�����"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][4] = ns["��⨪�"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][5] = rm["�⠯_��"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][6] = rm["��㤮��_��"] / 2.5
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][7] = rm["��㤮��_��"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][8] = rm["���_�_�"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][9] = rm["���_�_�"]
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][10] = 0
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][11] = 0

                      if rm["���_�_�"] < date_b & strlen(trim(rm["���_�_�"])) = 0 | rm["���_�_�"] > date_e
                      else
                        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_ࠡ_�""='" + str(rm["ID_ࠡ_�"],1,0) + "'")
                        kolvo_wu = sz["Kolvo"]
                        close(sz)
                        if kolvo_wu > 0 then
                          wu = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_ࠡ_�""='" + str(rm["ID_ࠡ_�"],1,0) + "'")
                          getfirst(wu)
                          stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][10] = wu["����_䠪�"]
                          stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][11] = wu["��㤮��_�"]
                          close(wu)
                        end if
                      end if
                    else
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][6] = stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][6] + rm["��㤮��_��"] / 2.5
                      stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][7] = stanki_to[ns["���_��ꥪ�"] + ns["��⨪�"] + ns["���_�����"]][7] + rm["��㤮��_��"]
                    end if
                  else
                    print "� �⠭�� � �������� ����஬ " + ns["���_�����"] + " �� 㪠��� ��. �⠭�� �� ����祭 � ����!"
                  end if
                end if
              end if
              getnext(rm)
            wend
            getnext(ns)
          wend
          close(ns)
          w_progress

          line = 7
          count = 0
          for i = 1 to 30
            itogi[i] = 0
          next
          sto = keyfirst(stanki_to)
          while( strlen(trim(sto)) != 0 )
            count = count + 1
            line = line + 1
            dde_poke("R" + str(line,1,0) + "C1",count)
            dde_poke("R" + str(line,1,0) + "C2","'" + trim(stanki_to[sto][1]))
            dde_poke("R" + str(line,1,0) + "C3","'" + trim(stanki_to[sto][2]))
            dde_poke("R" + str(line,1,0) + "C4",trim(stanki_to[sto][3]))
            dde_poke("R" + str(line,1,0) + "C5",trim(stanki_to[sto][4]))
            dde_poke("R" + str(line,1,0) + "C6",trim(stanki_to[sto][5]))
            dde_poke("R" + str(line,1,0) + "C7",stanki_to[sto][6])
            dde_poke("R" + str(line,1,0) + "C8",stanki_to[sto][7])
            if strlen(trim(stanki_to[sto][8])) != 0 then
              dde_poke("R" + str(line,1,0) + "C9","'" + substr(stanki_to[sto][8],7,2) + "." + substr(stanki_to[sto][8],5,2) + "." + substr(stanki_to[sto][8],1,4))
            end if
            if strlen(trim(stanki_to[sto][9])) != 0 then
              dde_poke("R" + str(line,1,0) + "C10","'" + substr(stanki_to[sto][9],7,2) + "." + substr(stanki_to[sto][9],5,2) + "." + substr(stanki_to[sto][9],1,4))
            end if
            dde_poke("R" + str(line,1,0) + "C11",stanki_to[sto][10])
            dde_poke("R" + str(line,1,0) + "C12",stanki_to[sto][11])
            itogi[7] = itogi[7] + stanki_to[sto][6]
            itogi[8] = itogi[8] + stanki_to[sto][7]
            itogi[11] = itogi[11] + stanki_to[sto][10]
            itogi[12] = itogi[12] + stanki_to[sto][11]
            sto = keynext(stanki_to,sto)
          wend

          if line >= 8 then
            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C6" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","�믮������ ࠡ�� � ࠬ��� ���")
            dde_poke("R" + str(line,1,0) + "C8",100)

            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C6" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","�⮣�")
            dde_poke("R" + str(line,1,0) + "C7",itogi[7])
            dde_poke("R" + str(line,1,0) + "C8",itogi[8] + 100)
            dde_poke("R" + str(line,1,0) + "C11",itogi[11])
            dde_poke("R" + str(line,1,0) + "C12",itogi[12])

            dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
            dde_execute("[BORDER(1,1,1,1,1)]")

            line = line + 2
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C4" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","��-0 �������筮� ���㦨�����")

            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C4" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","��-1 ���㣮����� ���㦨�����")

            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C4" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","��-2 ��������� ���㦨�����")

            line = line + 2
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C7" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C4","�㪮����⥫� ���ࠧ�������______________________________")
          end if
          dde_execute("[SELECT(" + chr(34) + "R1C1:R1C1" + chr(34) + ")]")
        else
          w_message("�訡��!","�� ������� �⠭��, ����� �ॡ���� �믮����� �� � ⥪�饬 �����!",0)
        end if

rem        print "SELECT * from OST_OBORUD where ""���_�����"" IS NOT NULL AND (""��_���_���"" IS NULL OR ""��_���_���""='1') ORDER BY ""ID_�����"" ASC"
        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from OST_OBORUD where ""���_�����"" IS NOT NULL AND ""��_���_���""='1' AND ""��������""='1'")
        kolvo = sz["Kolvo"]
        close(sz)

rem     ����� �� � �ᬮ��
        if kolvo > 0 then
          os = sql_query(Upper(getpath),"SELECT * from OST_OBORUD where ""���_�����"" IS NOT NULL AND ""��_���_���""='1' AND ""��������""='1' ORDER BY ""���_��ꥪ�"" ASC, ""��⨪�"" ASC")
          tecrem = keyarray
          osmotr = keyarray

          count_wp = 0
          w_progress("����� �ᬮ�� � ��...",0,kolvo)

          getfirst(os)
          while( eof(os) = 0 )
            count_wp = count_wp + 1
            w_progress(count_wp)
rem            print os["���_�����"]," // ",os["�����"]," // ",os["���_�����"]," // ",os["��������"]," // "

            id_oborud = os["ID_�����"]
            sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_OBORUD where ""ID_�����""='" + str(id_oborud,1,0) + "' AND SUBSTRING(""���_�_��"" FROM 1 FOR 4)='" + year + "'")
            kolvo_wo = sz["Kolvo"]
            close(sz)

            flag_osm = 0
            flag_rem = 0
            if kolvo_wo > 0
              rm = sql_query(Upper(getpath),"SELECT * from WORK_OBORUD where ""ID_�����""='" + str(id_oborud,1,0) + "' AND SUBSTRING(""���_�_��"" FROM 1 FOR 4)='" + year + "' ORDER BY ""ID_ࠡ_��"" ASC")
              getfirst(rm)
              while( eof(rm) = 0 & flag_osm = 0 & flag_rem = 0 )
                if substr(rm["���_�_��"],1,6) = year_month then
rem                  print rm["���_�_��"],rm["ID_�����"],rm["ID_ࠡ_��"],rm["ID_ࠡ_�"]
                  if strlen(trim(rm["�⠯_��"])) != 0 then
                    if strlen(trim(os["���_��ꥪ�"])) != 0
                      if find(vid,1,1,os["ID_��ઠ"]) = 0
                        if strpos("0",rm["�⠯_��"]) != 0
                          if keyfind(osmotr,str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]) = 0 then
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]] = array(10)
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][1] = os["���_��ꥪ�"]
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][2] = os["���_�����"]
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][3] = os["�����"]
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][4] = os["��⨪�"]
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][5] = '����'.vid * 0.4
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][6] = '����'.vid * 0.85 + '����'.vid * 0.2
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][7] = rm["���_�_�"]
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][8] = rm["���_�_�"]
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][9] = 0
                            osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][10] = 0

                            sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_ࠡ_�""='" + str(rm["ID_ࠡ_�"],1,0) + "'")
                            kolvo_wu = sz["Kolvo"]
                            close(sz)
                            if kolvo_wu > 0 then
                              wu = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_ࠡ_�""='" + str(rm["ID_ࠡ_�"],1,0) + "'")
                              getfirst(wu)
                              osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][9] = wu["����_䠪�"]
                              osmotr[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][10] = wu["��㤮��_�"]
                              close(wu)
                            end if
                          end if
                        else if strpos("1",rm["�⠯_��"]) != 0
                          if keyfind(tecrem,str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]) = 0 then
rem                            print "TP = " + os["���_�����"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]] = array(10)
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][1] = os["���_��ꥪ�"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][2] = os["���_�����"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][3] = os["�����"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][4] = os["��⨪�"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][5] = '����'.vid * 2
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][6] = '����'.vid * 6 + '����'.vid * 1.5
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][7] = rm["���_�_�"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][8] = rm["���_�_�"]
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][9] = 0
                            tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][10] = 0

                            sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_ࠡ_�""='" + str(rm["ID_ࠡ_�"],1,0) + "'")
                            kolvo_wu = sz["Kolvo"]
                            close(sz)
                            if kolvo_wu > 0 then
                              wu = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_ࠡ_�""='" + str(rm["ID_ࠡ_�"],1,0) + "'")
                              getfirst(wu)
                              tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][9] = wu["����_䠪�"]
                              tecrem[str(rm["ID_ࠡ_�"],1,0) + os["���_�����"]][10] = wu["��㤮��_�"]
                              close(wu)
                            end if
                          end if
                        else
                          print "�������⭮� �� ",rm["�⠯_��"],os["���_�����"]
                        end if
                      else
                        print "�訡��! �� ������� ID_��ઠ ",os["ID_��ઠ"]
                      end if
                    else
                      print "� �⠭�� � �������� ����஬ " + os["���_�����"] + " �� 㪠��� ��. �⠭�� �� ����祭 � ����!"
                    end if
                  end if
                end if
                getnext(rm)
              wend
            else
              print "�� ������� ����� � WORK_OBORUD ",id_oborud
            end if
            close(rm)
            getnext(os)
          wend
          close(os)
          w_progress
        end if

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "�ᬮ��" + chr(34) + ")]")
        dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",9,FALSE,FALSE,FALSE,FALSE,1)]")
        dde_execute("[CLEAR()]")
        dde_execute("[BORDER(0,0,0,0,0)]")
        dde_poke("R2C1","�⤥�� �28, ᮣ��᭮ �������� ��䨪� ��� � " + name_months[month][2] + " ����� " + year + " �.")
        line = 7
        count = 0
        for i = 1 to 30
          itogi[i] = 0
        next
        s = keyfirst(osmotr)
        while( strlen(trim(s)) != 0 )
          line = line + 1
          count = count + 1
          dde_poke("R" + str(line,1,0) + "C1",count)
          dde_poke("R" + str(line,1,0) + "C2","'" + trim(osmotr[s][1]))
          dde_poke("R" + str(line,1,0) + "C3","'" + trim(osmotr[s][2]))
          dde_poke("R" + str(line,1,0) + "C4",trim(osmotr[s][3]))
          dde_poke("R" + str(line,1,0) + "C5",trim(osmotr[s][4]))
          dde_poke("R" + str(line,1,0) + "C6","�")
          dde_poke("R" + str(line,1,0) + "C7",osmotr[s][5])
          dde_poke("R" + str(line,1,0) + "C8",osmotr[s][6])
          if strlen(trim(osmotr[s][7])) != 0 then
            dde_poke("R" + str(line,1,0) + "C9","'" + substr(osmotr[s][7],7,2) + "." + substr(osmotr[s][7],5,2) + "." + substr(osmotr[s][7],1,4))
          end if
          if strlen(trim(osmotr[s][8])) != 0 then
            dde_poke("R" + str(line,1,0) + "C10","'" + substr(osmotr[s][8],7,2) + "." + substr(osmotr[s][8],5,2) + "." + substr(osmotr[s][8],1,4))
          end if
          if osmotr[s][7] < date_b & strlen(trim(osmotr[s][8])) = 0 | osmotr[s][8] > date_e
          else
            dde_poke("R" + str(line,1,0) + "C11",osmotr[s][9])
            dde_poke("R" + str(line,1,0) + "C12",osmotr[s][10])
            itogi[3] = itogi[3] + osmotr[s][9]
            itogi[4] = itogi[4] + osmotr[s][10]
          end if

          itogi[1] = itogi[1] + osmotr[s][5]
          itogi[2] = itogi[2] + osmotr[s][6]
          s = keynext(osmotr,s)
        wend
        if line >= 8 then
          line = line + 1
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C6" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C1","�⮣�")
          dde_poke("R" + str(line,1,0) + "C7",itogi[1])
          dde_poke("R" + str(line,1,0) + "C8",itogi[2])
          dde_poke("R" + str(line,1,0) + "C11",itogi[3])
          dde_poke("R" + str(line,1,0) + "C12",itogi[4])

          dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
          dde_execute("[BORDER(1,1,1,1,1)]")

          line = line + 3
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C6" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C4","�㪮����⥫� ���ࠧ�������______________________________")
        end if
        dde_execute("[SELECT(" + chr(34) + "R1C1:R1C1" + chr(34) + ")]")

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "��" + chr(34) + ")]")
        dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",9,FALSE,FALSE,FALSE,FALSE,1)]")
        dde_execute("[CLEAR()]")
        dde_execute("[BORDER(0,0,0,0,0)]")
        dde_poke("R2C1",name_months[month][2] + " ����� " + year + " �. �⤥� �28")
        line = 7
        count = 0
        for i = 1 to 30
          itogi[i] = 0
        next
        s = keyfirst(tecrem)
        while( strlen(trim(s)) != 0 )
          line = line + 1
          count = count + 1
          dde_poke("R" + str(line,1,0) + "C1",count)
          dde_poke("R" + str(line,1,0) + "C2","'" + trim(tecrem[s][1]))
          dde_poke("R" + str(line,1,0) + "C3","'" + trim(tecrem[s][2]))
          dde_poke("R" + str(line,1,0) + "C4",trim(tecrem[s][3]))
          dde_poke("R" + str(line,1,0) + "C5",trim(tecrem[s][4]))
          dde_poke("R" + str(line,1,0) + "C6","��")
          dde_poke("R" + str(line,1,0) + "C7",tecrem[s][5])
          dde_poke("R" + str(line,1,0) + "C8",tecrem[s][6])

          if strlen(trim(tecrem[s][7])) != 0 then
            dde_poke("R" + str(line,1,0) + "C9","'" + substr(tecrem[s][7],7,2) + "." + substr(tecrem[s][7],5,2) + "." + substr(tecrem[s][7],1,4))
          end if
          if strlen(trim(tecrem[s][8])) != 0 then
            dde_poke("R" + str(line,1,0) + "C10","'" + substr(tecrem[s][8],7,2) + "." + substr(tecrem[s][8],5,2) + "." + substr(tecrem[s][8],1,4))
          end if
          if tecrem[s][7] < date_b & strlen(trim(tecrem[s][8])) = 0 | tecrem[s][8] > date_e
          else
            dde_poke("R" + str(line,1,0) + "C11",tecrem[s][9])
            dde_poke("R" + str(line,1,0) + "C12",tecrem[s][10])
            itogi[3] = itogi[3] + tecrem[s][9]
            itogi[4] = itogi[4] + tecrem[s][10]
          end if
          itogi[1] = itogi[1] + tecrem[s][5]
          itogi[2] = itogi[2] + tecrem[s][6]
          s = keynext(tecrem,s)
        wend
        if line >= 8 then
          line = line + 1
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C6" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C1","�⮣�")
          dde_poke("R" + str(line,1,0) + "C7",itogi[1])
          dde_poke("R" + str(line,1,0) + "C8",itogi[2])
          dde_poke("R" + str(line,1,0) + "C11",itogi[3])
          dde_poke("R" + str(line,1,0) + "C12",itogi[4])

          dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
          dde_execute("[BORDER(1,1,1,1,1)]")

          line = line + 3
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C7" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C4","�㪮����⥫� ���ࠧ�������______________________________")
        end if
        dde_execute("[SELECT(" + chr(34) + "R1C1:R1C1" + chr(34) + ")]")

        pf = array(2)
        for i = 1 to 2
          pf[i] = array(14)
          for j = 1 to 14
            pf[i][j] = 0
          next
        next

        sto = keyfirst(stanki_to)
        while( strlen(trim(sto)) != 0 )
          if keyfind(stops,stanki_to[sto][1]) = 0 then
            stops[stanki_to[sto][1]] = array(40)
            for i = 1 to 40
              stops[stanki_to[sto][1]][i] = 0
            next
          end if
          stops[stanki_to[sto][1]][5] = stops[stanki_to[sto][1]][5] + 1
          stops[stanki_to[sto][1]][6] = stops[stanki_to[sto][1]][6] + stanki_to[sto][6]
rem          stops[stanki_to[sto][1]][7] = stops[stanki_to[sto][1]][7] + stanki_to[sto][7]
          stops[stanki_to[sto][1]][11] = stops[stanki_to[sto][1]][11] + stanki_to[sto][6]

          stops[stanki_to[sto][1]][16] = stops[stanki_to[sto][1]][16] + 1
          stops[stanki_to[sto][1]][17] = stops[stanki_to[sto][1]][17] + stanki_to[sto][10]
          stops[stanki_to[sto][1]][22] = stops[stanki_to[sto][1]][22] + stanki_to[sto][10]

          pf[1][3] = pf[1][3] + stanki_to[sto][7]
          pf[2][3] = pf[2][3] + stanki_to[sto][10]

          sto = keynext(stanki_to,sto)
        wend

        sto = keyfirst(tecrem)
        while( strlen(trim(sto)) != 0 )
          if keyfind(stops,tecrem[sto][1]) = 0 then
            stops[tecrem[sto][1]] = array(40)
            for i = 1 to 40
              stops[tecrem[sto][1]][i] = 0
            next
          end if
          stops[tecrem[sto][1]][1] = stops[tecrem[sto][1]][1] + 1
          stops[tecrem[sto][1]][2] = stops[tecrem[sto][1]][2] + tecrem[sto][5]
          stops[tecrem[sto][1]][11] = stops[tecrem[sto][1]][11] + tecrem[sto][5]

          stops[tecrem[sto][1]][12] = stops[tecrem[sto][1]][12] + 1
          stops[tecrem[sto][1]][13] = stops[tecrem[sto][1]][13] + tecrem[sto][9]
          stops[tecrem[sto][1]][22] = stops[tecrem[sto][1]][22] + tecrem[sto][9]

          pf[1][1] = pf[1][1] + tecrem[sto][6]
          pf[2][1] = pf[2][1] + tecrem[sto][10]
          sto = keynext(tecrem,sto)
        wend

        sto = keyfirst(osmotr)
        while( strlen(trim(sto)) != 0 )
          if keyfind(stops,osmotr[sto][1]) = 0 then
            stops[osmotr[sto][1]] = array(40)
            for i = 1 to 40
              stops[osmotr[sto][1]][i] = 0
            next
          end if
          stops[osmotr[sto][1]][3] = stops[osmotr[sto][1]][3] + 1
          stops[osmotr[sto][1]][4] = stops[osmotr[sto][1]][4] + osmotr[sto][5]
          stops[osmotr[sto][1]][11] = stops[osmotr[sto][1]][11] + osmotr[sto][5]

          stops[osmotr[sto][1]][14] = stops[osmotr[sto][1]][14] + 1
          stops[osmotr[sto][1]][15] = stops[osmotr[sto][1]][15] + osmotr[sto][9]
          stops[osmotr[sto][1]][22] = stops[osmotr[sto][1]][22] + osmotr[sto][9]

          pf[1][2] = pf[1][2] + osmotr[sto][6]
          pf[2][2] = pf[2][2] + osmotr[sto][10]
          sto = keynext(osmotr,sto)
        wend

        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from OST_OBORUD where ""���_�����"" IS NOT NULL")
        kolvo = sz["Kolvo"]
        close(sz)
        if kolvo > 0 then
          all = sql_query(Upper(getpath),"SELECT * from OST_OBORUD where ""���_�����"" IS NOT NULL")
          if flag_two = 0
            yo1 = open("Y" + year + "_oborud",-2)
          else
            yo1 = open("Y" + substr(date_start,1,4) + "_oborud",-2)
            yo2 = open("Y" + substr(date_finish,1,4) + "_oborud",-2)
          end if

          subgroups = keyarray
          kol_oborud = 0
          getfirst(all)
          while( eof(all) = 0 )
            flag_sg = 0
            kol_oborud = kol_oborud + 1
            if keyfind(stops,all["���_��ꥪ�"]) = 0
              stops[all["���_��ꥪ�"]] = array(40)
              for i = 1 to 40
                stops[all["���_��ꥪ�"]][i] = 0
              next
              stops[all["���_��ꥪ�"]][23] = stops[all["���_��ꥪ�"]][23] + 1
            else
              stops[all["���_��ꥪ�"]][23] = stops[all["���_��ꥪ�"]][23] + 1
            end if
            if find(vid,1,1,all["ID_��ઠ"]) = 0
              if strlen(trim('�����㯯�'.vid)) != 0
                flag_sg = 1
                if keyfind(subgroups,trim('�����㯯�'.vid)) = 0
                  subgroups[trim('�����㯯�'.vid)] = array(5)
                  subgroups[trim('�����㯯�'.vid)][1] = 1
                  subgroups[trim('�����㯯�'.vid)][2] = 0
                  subgroups[trim('�����㯯�'.vid)][3] = 0
                else
                  subgroups[trim('�����㯯�'.vid)][1] = subgroups[trim('�����㯯�'.vid)][1] + 1
                end if
              else
                print "34 // ",all["���_�����"]," ��������� �������� �����㯯�"
              end if
            else
              print "33"
            end if

            if flag_two = 0
              if findge(yo1,2,1,all["ID_�����"],date_start) = 0 then
                while( eof(yo1) = 0 & all["ID_�����"] = 'ID_�����'.yo1 & '���'.yo1 <= date_finish )
                  stops[all["���_��ꥪ�"]][24] = stops[all["���_��ꥪ�"]][24] + '��ࠡ_����'.yo1
                  if flag_sg = 1 then
                    subgroups[trim('�����㯯�'.vid)][2] = subgroups[trim('�����㯯�'.vid)][2] + '��ࠡ_����'.yo1
                  end if
                  getnext(yo1)
                wend
              end if
            else
              if findge(yo1,2,1,all["ID_�����"],date_start) = 0 then
                while( eof(yo1) = 0 & all["ID_�����"] = 'ID_�����'.yo1 & '���'.yo1 <= date_finish )
                  stops[all["���_��ꥪ�"]][24] = stops[all["���_��ꥪ�"]][24] + '��ࠡ_����'.yo1
                  if flag_sg = 1 then
                    subgroups[trim('�����㯯�'.vid)][2] = subgroups[trim('�����㯯�'.vid)][2] + '��ࠡ_����'.yo1
                  end if
                  getnext(yo1)
                wend
              end if
              if findge(yo2,2,1,all["ID_�����"],date_start) = 0 then
                while( eof(yo2) = 0 & all["ID_�����"] = 'ID_�����'.yo2 & '���'.yo2 <= date_finish )
                  stops[all["���_��ꥪ�"]][24] = stops[all["���_��ꥪ�"]][24] + '��ࠡ_����'.yo2
                  if flag_sg = 1 then
                    subgroups[trim('�����㯯�'.vid)][2] = subgroups[trim('�����㯯�'.vid)][2] + '��ࠡ_����'.yo2
                  end if
                  getnext(yo2)
                wend
              end if
            end if
            getnext(all)
          wend
          close(all)
        end if

rem     ������ ࠡ�祣� �६���
        common_fond = 0
        otpuska = 0
        blist = 0
        proch_lost = 0
        tmr = 0
        kolsotr = 0
        als = open("all_sotr")
        grvs = open("grafic_rv_sotrud")
        typetab = open("type_tabel")

rem        print "SELECT * from TABELRV where ""���"">='" + date_b + "' AND ""���""<='" + date_e + "' ORDER BY ""ID_���,���"" ASC"

        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from TABELRV where ""���"">='" + date_b + "' AND ""���""<='" + date_e + "'")
        kolvo = sz["Kolvo"]
        close(sz)
        if kolvo > 0 then
          trv = sql_query(Upper(getpath),"SELECT * from TABELRV where ""���"">='" + date_b + "' AND ""���""<='" + date_e + "' ORDER BY ""ID_���"" ASC, ""���"" ASC")
          getfirst(trv)
          while( eof(trv) = 0 )
            if find(als,1,1,'ID_���'.trv) = 0
              if '�������'.als = 0
                print '����㤭��'.als,"\\",'���'.trv,"\\",'����_���'.trv
                if keyfind(people,str('ID_���'.als,1,0) + '����㤭��'.als) = 0 then
                  people[str('ID_���'.als,1,0) + '����㤭��'.als] = 0
                end if
                if strpos("��������",Upper('����_���'.trv)) != 0
                else
                  if find(grvs,1,1,'ID_�������'.trv) = 0
                    fl_smena = 0
                    for i = 1 to 6
                      field_start = "���_" + str(i,1,0)
                      field_finish = "���_" + str(i,1,0)
                      if strlen(trim('field_start'.grvs)) != 0 & strlen(trim('field_finish'.grvs)) != 0 then
                        if fl_smena = 0 then
                          fl_smena = 1
                        end if
                      end if
                    next
                    if fl_smena = 1 then
                      for i = 1 to 6
                        field_start = "���_" + str(i,1,0)
                        field_finish = "���_" + str(i,1,0)
                        if strlen(trim('field_start'.trv)) != 0 & strlen(trim('field_finish'.trv)) != 0 then
                          common_fond = common_fond + kol_chasov('field_start'.trv,'field_finish'.trv)
                          people[str('ID_���'.als,1,0) + '����㤭��'.als] = people[str('ID_���'.als,1,0) + '����㤭��'.als] + kol_chasov('field_start'.trv,'field_finish'.trv)
                          if strpos("28-6",'�ਣ���'.als) != 0 then
                            pf[1][4] = pf[1][4] + kol_chasov('field_start'.trv,'field_finish'.trv)
                          end if
                        end if
                      next
                    end if
                  else if find(typetab,1,1,'ID_�������'.trv) = 0
                    if '����'.typetab = 1
  rem                    print "���� ",otpuska
                      if '����_横�'.trv = 1 | '����_横�'.trv = 2
                        otpuska = otpuska + 11
                        people[str('ID_���'.als,1,0) + '����㤭��'.als] = people[str('ID_���'.als,1,0) + '����㤭��'.als] + 11
                      else if '����_横�'.trv = 0 & trim('���_���'.trv) != "�㡡��" & trim('���_���'.trv) != "����ᥭ�"
                        otpuska = otpuska + 8
                        people[str('ID_���'.als,1,0) + '����㤭��'.als] = people[str('ID_���'.als,1,0) + '����㤭��'.als] + 8
                      end if
rem                      print "���� ",otpuska
                    else if strpos("����������",Upper('����_⠡'.typetab)) != 0
                      if '����_横�'.trv = 1 | '����_横�'.trv = 2
                        blist = blist + 11
                        people[str('ID_���'.als,1,0) + '����㤭��'.als] = people[str('ID_���'.als,1,0) + '����㤭��'.als] + 11
                      else if '����_横�'.trv = 0 & trim('���_���'.trv) != "�㡡��" & trim('���_���'.trv) != "����ᥭ�"
                        blist = blist + 8
                        people[str('ID_���'.als,1,0) + '����㤭��'.als] = people[str('ID_���'.als,1,0) + '����㤭��'.als] + 8
                      end if
                    else
                      proch_lost = proch_lost + 8
                    end if
                  end if
                end if
              else
              end if
            else
              w_message("�訡��!","�� ������ ���㤭�� � �⤥�� ���஢ " + trim('����㤭��'.trv) + "!",0)
            end if
            getnext(trv)
          wend
          close(trv)
        end if

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "���" + chr(34) + ")]")
        dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",12,FALSE,FALSE,FALSE,FALSE,1)]")
        dde_execute("[CLEAR()]")
        dde_execute("[BORDER(0,0,0,0,0)]")

        dde_poke("R3C1","�⤥�� � 28, � " + name_months[month][2] + " ����� " + year + " �.")

        line = 8
        dde_poke("R8C6","�����⠦ � ����⠭���� ��ண� ����㤮�����")
        dde_poke("R8C8",num(str(pf[1][4] / 2,1,0)))
        dde_poke("R9C6","�ਥ��� ������ ����㤮�����")
        dde_poke("R9C8",pf[1][4] - num(str(pf[1][4] / 2,1,0)))
        line = line + 3
        dde_poke("R" + str(line,1,0) + "C8",pf[1][4])
        dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[BORDER(1,1,1,1,1)]")

        line = line + 2
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C6" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C4","�㪮����⥫� ���ࠧ�������______________________________")

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "��" + chr(34) + ")]")
        dde_poke("R3C1","६��� ����㤮����� � " + name_months[month][2] + " ����� " + year + " �.")
        line = 8
rem        print "SELECT * from WORK_UNIT where ( ""ID_�����"" IS NOT NULL OR ""ID_�����"">'0' ) AND ( ( ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' AND ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' ) OR ( ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' AND ( ""���_�_�"">'" + date_e + "' OR ""���_�_�"" IS NULL ) ) OR ( ""���_�_�""<='" + date_b + "' AND ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "') OR ( ""���_�_�""<'" + date_b + "' AND ( ""���_�_�"" IS NULL OR ""���_�_�"">'" + date_e + "') ) ) ORDER BY ""���_�_�"" ASC"

        vr = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_�����"">'0' AND ( ( ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' AND ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' ) OR ( ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' AND ( ""���_�_�"">'" + date_e + "' OR ""���_�_�"" IS NULL ) ) OR ( ""���_�_�""<='" + date_b + "' AND ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "') OR ( ""���_�_�""<'" + date_b + "' AND ( ""���_�_�"" IS NULL OR ""���_�_�"">'" + date_e + "') ) ) ")
        kolvo = vr["Kolvo"]
        close(vr)

        if kolvo > 0 then
          count = 0

          for i = 1 to 30
            itogi[i] = 0
          next
          ost = open("ost_oborud")
          vr = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_�����"">'0' AND ( ( ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' AND ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' ) OR ( ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "' AND ( ""���_�_�"">'" + date_e + "' OR ""���_�_�"" IS NULL ) ) OR ( ""���_�_�""<='" + date_b + "' AND ""���_�_�"">='" + date_b + "' AND ""���_�_�""<='" + date_e + "') OR ( ""���_�_�""<'" + date_b + "' AND ( ""���_�_�"" IS NULL OR ""���_�_�"">'" + date_e + "') ) ) ORDER BY ""���_�_�"" ASC")
          getfirst(vr)
          while( eof(vr) = 0 )
rem            print vr["ID_ࠡ_�"]
            if findge(ost,1,3,vr["ID_㧫�"]) = 0 then
              while( eof(ost) = 0 & vr["ID_㧫�"] = 'ID_㧫�'.ost )
                line = line + 1
                count = count + 1
                dde_poke("R" + str(line,1,0) + "C1","'" + str(count,1,0))
                dde_poke("R" + str(line,1,0) + "C2","'" + trim('���_��ꥪ�'.ost))
                dde_poke("R" + str(line,1,0) + "C3","'" + trim('���_�����'.ost))
rem                print trim('���_�����'.ost)
                dde_poke("R" + str(line,1,0) + "C4",trim('�����'.ost))
                dde_poke("R" + str(line,1,0) + "C5","'" + trim('��⨪�'.ost))
rem                dde_poke("R" + str(line,1,0) + "C6",trim(vr["�ਬ_�_�"]))

                sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_OBORUD where ""ID_ࠡ_�""='" + str(vr["ID_ࠡ_�"],1,0) + "'")
                kolvo_wo = sz["Kolvo"]
                close(sz)
                if kolvo_wo > 0 then
                  wo = sql_query(Upper(getpath),"SELECT * from WORK_OBORUD where ""ID_ࠡ_�""='" + str(vr["ID_ࠡ_�"],1,0) + "' ORDER BY ""ID_ࠡ_��"" ASC")
                  list_work = keyarray
                  getfirst(wo)
                  while( eof(wo) = 0 )
                    if keyfind(list_work,str(wo["ID_ࠡ_��"],1,0)) = 0 then
                      list_work[str(wo["ID_ࠡ_��"],1,0)] = array(3)
                      list_work[str(wo["ID_ࠡ_��"],1,0)][1] = trim(wo["�ਬ_�_�"])
                      list_work[str(wo["ID_ࠡ_��"],1,0)][2] = trim(wo["���_ࠡ��"])
                      list_work[str(wo["ID_ࠡ_��"],1,0)][3] = trim(wo["���ᒥ孮�"])
                    end if
                    getnext(wo)
                  wend
                  close(wo)
                end if
                lw = ""
                s = keyfirst(list_work)
                while( strlen(trim(s)) != 0 )
                  lw = lw + trim(list_work[s][2])
                  s = keynext(list_work,s)
                wend
                dde_poke("R" + str(line,1,0) + "C6",lw)
rem                print vr["���_�_�"]," // ",date_b," // ",strlen(trim(vr["���_�_�"]))," // ",vr["���_�_�"]," // ",date_e,"//"
                if ( vr["���_�_�"] < substr(date_b,7,4) + substr(date_b,4,2) + substr(date_b,1,2) & strlen(trim(vr["���_�_�"])) = 0 ) | ( vr["���_�_�"] > substr(date_e,7,4) + substr(date_e,4,2) + substr(date_e,1,2) )
                else
rem print "11"
                  dde_poke("R" + str(line,1,0) + "C7",vr["����_����"])
                  dde_poke("R" + str(line,1,0) + "C8",vr["��㤮��_��"])
                  if substr(date_b,7,4) = substr(vr["���_�_�"],1,4) | strlen(trim(vr["���_�_�"])) = 0
                    tablname = "Y" + substr(date_b,7,4) + "_OBORUD"
rem                    print tablname,"// ","SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4) + "'"
                    if strlen(trim(vr["���_�_�"])) = 0
                      kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='" + date_e + "'")
                    else
                      kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4) + "'")
                    end if
                    kolyobor = kol_yo["Kolvo"]
                    close(kol_yo)

                    if kolyobor > 0 then
                      if strlen(trim(vr["���_�_�"])) = 0
                        yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='" + date_e + "'")
                      else
                        yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4) + "'")
                      end if
                      prostoi = 0
                      getfirst(yobor)
                      while( eof(yobor) = 0 )
                        if vr["���_�_�"] != yobor["���"]
                          prostoi = prostoi + yobor["���_����"]
                        else
                          if vr["���_�_�"] = vr["���_�_�"]
                            prostoi = prostoi + ( num(substr(vr["�६�_�_�"],1,2)) * 60 + num(substr(vr["�६�_�_�"],4,2)) - ( num(substr(vr["�६�_�_�"],1,2)) * 60 + num(substr(vr["�६�_�_�"],4,2)) ) ) / 60
                          else
                            prostoi = prostoi + (( num(substr(vr["�६�_�_�"],1,2)) - 8 ) * 60 + num(substr(vr["�६�_�_�"],4,2))) / 60
                          end if
                        end if
                        getnext(yobor)
                      wend
                      close(yobor)
                    end if
                  else
                    prostoi = 0
                    tablname = "Y" + substr(date_b,7,4) + "_OBORUD"
rem                    print tablname,"// ","SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='31.12." + substr(date_b,7,4) + "'"
                    kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='31.12." + substr(date_b,7,4) + "'")
                    kolyobor = kol_yo["Kolvo"]
                    close(kol_yo)

                    if kolyobor > 0 then
                      yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='" + date_b + "' AND ""���""<='31.12." + substr(date_b,7,4) + "'")
                      getfirst(yobor)
                      while( eof(yobor) = 0 )
                        if vr["���_�_�"] != yobor["���"]
                          prostoi = prostoi + yobor["���_����"]
                        else
                          if vr["���_�_�"] = vr["���_�_�"]
                            prostoi = prostoi + ( num(substr(vr["�६�_�_�"],1,2)) * 60 + num(substr(vr["�६�_�_�"],4,2)) - ( num(substr(vr["�६�_�_�"],1,2)) * 60 + num(substr(vr["�६�_�_�"],4,2)) ) ) / 60
                          else
                            prostoi = prostoi + (( num(substr(vr["�६�_�_�"],1,2)) - 8 ) * 60 + num(substr(vr["�६�_�_�"],4,2))) / 60
                          end if
                        end if
                        getnext(yobor)
                      wend
                      close(yobor)
                    end if

                    tablname = "Y" + substr(vr["���_�_�"],1,4) + "_OBORUD"
rem                    print tablname,"// ","SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='01.01." + substr(vr["���_�_�"],1,4) + "' AND ""���""<='" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4)
                    kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='01.01." + substr(vr["���_�_�"],1,4) + "' AND ""���""<='" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4))
                    kolyobor = kol_yo["Kolvo"]
                    close(kol_yo)

                    if kolyobor > 0 then
                      yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_�����""='" + str('ID_�����'.ost,1,0) + "' AND ""���"">='01.01." + substr(vr["���_�_�"],1,4) + "' AND ""���""<='" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4))
                      getfirst(yobor)
                      while( eof(yobor) = 0 )
                        if vr["���_�_�"] != yobor["���"]
                          prostoi = prostoi + yobor["���_����"]
                        else
                          if vr["���_�_�"] = vr["���_�_�"]
                            prostoi = prostoi + ( num(substr(vr["�६�_�_�"],1,2)) * 60 + num(substr(vr["�६�_�_�"],4,2)) - ( num(substr(vr["�६�_�_�"],1,2)) * 60 + num(substr(vr["�६�_�_�"],4,2)) ) ) / 60
                          else
                            prostoi = prostoi + (( num(substr(vr["�६�_�_�"],1,2)) - 8 ) * 60 + num(substr(vr["�६�_�_�"],4,2))) / 60
                          end if
                        end if
                        getnext(yobor)
                      wend
                      close(yobor)
                    end if
                  end if
                  dde_poke("R" + str(line,1,0) + "C11",prostoi)

rem                  dde_poke("R" + str(line,1,0) + "C11",vr["����_䠪�"])
                  dde_poke("R" + str(line,1,0) + "C12",vr["��㤮��_�"])
                  itogi[1] = itogi[1] + vr["����_����"]
                  itogi[2] = itogi[2] + vr["��㤮��_��"]
                  itogi[3] = itogi[3] + vr["����_䠪�"]
                  itogi[4] = itogi[4] + vr["��㤮��_�"]
                end if

                if strlen(trim(vr["���_�_�"])) != 0 then
                  dde_poke("R" + str(line,1,0) + "C9","'" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4))
                end if
                if strlen(trim(vr["���_�_�"])) != 0 then
                  dde_poke("R" + str(line,1,0) + "C10","'" + substr(vr["���_�_�"],7,2) + "." + substr(vr["���_�_�"],5,2) + "." + substr(vr["���_�_�"],1,4))
                end if
                if find(vid,1,1,'ID_��ઠ'.ost) = 0 then
                  if strlen(trim('�����㯯�'.vid)) != 0
                    if keyfind(subgroups,trim('�����㯯�'.vid)) = 0
                      subgroups[trim('�����㯯�'.vid)] = array(5)
                      subgroups[trim('�����㯯�'.vid)][1] = 1
                      subgroups[trim('�����㯯�'.vid)][2] = 0
                      subgroups[trim('�����㯯�'.vid)][3] = 1
                    else
                      subgroups[trim('�����㯯�'.vid)][3] = subgroups[trim('�����㯯�'.vid)][3] + 1
                    end if
                  else 
                    print "34-1 // ",'���_�����'.ost," ��������� �������� �����㯯�"
                  end if
                end if  
                getnext(ost)
              wend
            end if
            getnext(vr)
          wend
          close(vr)
          line = line + 20
          dde_poke("R" + str(line,1,0) + "C3","�����")
          dde_poke("R" + str(line,1,0) + "C7",itogi[1])
          dde_poke("R" + str(line,1,0) + "C8",itogi[2])
          dde_poke("R" + str(line,1,0) + "C11",itogi[3])
          dde_poke("R" + str(line,1,0) + "C12",itogi[4])

          dde_execute("[SELECT(" + chr(34) + "R9C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
          dde_execute("[BORDER(1,1,1,1,1)]")
        end if

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "������" + chr(34) + ")]")
        dde_poke("R2C1","��� �믮������ ࠡ�� �� " + name_months[month][1] + " " + year + " �., �/��.")
        dde_poke("R9C1","���⮩ �孮�����᪮�� ����㤮�����, ��.")
rem  ����� ᤥ����

        stroka = ""
        if r = 6
          if flag_insert = 1
            dde_poke("R5C7",pf[1][1])
            dde_poke("R5C8",pf[1][2])
            dde_poke("R5C9",pf[1][3] + 100)
            dde_poke("R5C10",pf[1][4])
            dde_poke("R5C11",itogi[2])
            dde_poke("R5C13",pf[1][1] + pf[1][2] + pf[1][3] + 100 + pf[1][4] + itogi[2])
            dde_poke("R5C14",common_fond)
            dde_poke("R5C15",otpuska)
            dde_poke("R5C16",blist)
            dde_poke("R5C17",proch_lost)
            dde_poke("R5C18",otpuska + blist + proch_lost)
            stroka = stroka + "R5C11:" + str(itogi[2],1,2) + ";" + "R5C14:" + str(common_fond,1,2) + ";" + "R5C15:" + str(otpuska,1,2) + ";" + "R5C16:" + str(blist,1,2) + ";" + "R5C17:" + str(proch_lost,1,2) + ";" + "R5C18:" + str(otpuska + blist + proch_lost,1,2) + ";"
          else
            pos = keyfirst(svod_plan)
            while( strlen(trim(pos)) != 0 )
              print pos," = ",svod_plan[pos]
              dde_poke(pos,svod_plan[pos])
              pos = keynext(svod_plan,pos)
            wend
          end if
        else if r = 7
          if flag_pos = 1 then
            pos = keyfirst(svod_plan)
            while( strlen(trim(pos)) != 0 )
              dde_poke(pos,svod_plan[pos])
              pos = keynext(svod_plan,pos)
            wend
          end if
          dde_poke("R5C7",pf[1][1])
          dde_poke("R5C8",pf[1][2])
          dde_poke("R5C9",pf[1][3])
          dde_poke("R5C10",pf[1][4])
          dde_poke("R5C11",itogi[2])
          dde_poke("R5C13",pf[1][1] + pf[1][2] + pf[1][3] + 100 + pf[1][4] + itogi[2])
          dde_poke("R5C14",common_fond)
          dde_poke("R5C15",otpuska)
          dde_poke("R5C16",blist)
          dde_poke("R5C17",proch_lost)
          dde_poke("R5C18",otpuska + blist + proch_lost)
          dde_poke("R6C7",pf[2][1])
          dde_poke("R6C8",pf[2][2])
          dde_poke("R6C9",pf[2][3])
          dde_poke("R6C11",itogi[4])
          dde_poke("R6C13",pf[2][1] + pf[2][2] + pf[2][3] + 100 + pf[2][4] + itogi[4])
        end if

        line = 14
        for i = 1 to 30
          itogi[i] = 0
        next

        s = keyfirst(stops)
        while( strlen(trim(s)) != 0 )
          line = line + 1
          print stops[s][1],stops[s][2],stops[s][3],stops[s][4],stops[s][5],stops[s][6],stops[s][7]
          dde_poke("R" + str(line,1,0) + "C1","'" + trim(s))
          for i = 1 to 22
            if i != 11 & i != 22
              dde_poke("R" + str(line,1,0) + "C" + str(i + 1,1,0),stops[s][i])
            else
              dde_poke("R" + str(line,1,0) + "C" + str(i + 1,1,0),"=RC[-9]+RC[-7]+RC[-5]+RC[-3]+RC[-1]")
            end if
            itogi[i] = itogi[i] + stops[s][i]
          next
          dde_poke("R" + str(line,1,0) + "C24",stops[s][23])
          dde_poke("R" + str(line,1,0) + "C25",stops[s][24])
          dde_poke("R" + str(line,1,0) + "C27","=RC[-4]/RC[-2]*100")
          itogi[23] = itogi[23] + stops[s][23]
          itogi[24] = itogi[24] + stops[s][24]
          s = keynext(stops,s)
        wend

        line = line + 1
        dde_poke("R" + str(line,1,0) + "C1","�⮣�")
        for i = 1 to 22
          dde_poke("R" + str(line,1,0) + "C" + str(i + 1,1,0),itogi[i])
        next
        dde_poke("R14C24",itogi[23])
        dde_poke("R14C25",itogi[24])
        dde_poke("R" + str(line,1,0) + "C24",itogi[23])
        dde_poke("R" + str(line,1,0) + "C25",itogi[24])
        dde_poke("R" + str(line,1,0) + "C27","=RC[-4]/RC[-2]*100")
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C57" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT("+chr(34)+"Times New Roman"+chr(34)+",12,TRUE,FALSE,FALSE,FALSE,1)]")

        dde_execute("[SELECT(" + chr(34) + "R15C1:R" + str(line,1,0) + "C27" + chr(34) + ")]")
        dde_execute("[BORDER(1,1,1,1,1)]")

        if r = 6
          if flag_insert = 1
            stroka = stroka + "R5C7:" + str(pf[1][1],1,2) + ";" + "R5C8:" + str(pf[1][2],1,2) + ";" + "R5C9:" + str(pf[1][3] + 100,1,2) + ";" + "R5C10:" + str(pf[1][4],1,2) + ";" + "R5C13:" + str(pf[1][1] + pf[1][2] + pf[1][3] + 100 + pf[1][4],1,2) + ";"
            if find(po,2,2,year,month) != 0
              clearbuf(po)
              'ID'.po = id_po
              '���'.po = year
              '�����'.po = month
              '���祭��'.po = stroka
              insert(po)
            else
              '���'.po = year
              '�����'.po = month
              '���祭��'.po = stroka
              update(po)
            end if
          else
          end if
        else if r = 7
          if flag_pos = 1
          else
          end if
        end if

        line = line + 1
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C11" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C1","�����⨬� �����")
        dde_poke("R" + str(line,1,0) + "C12",1333.0)
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[BORDER(1,1,1,1,1)]")

        line = line + 3
        start_line = line
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C1","���������� ���⮩ ����㤮����� �� ��㯯��")

        line = line + 1
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C7" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C4","��饥 ������⢮")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C8:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C8","������⢮ ������ � ���⮥")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line + 2,1,0) + "C3" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C1","��㯯� ����㤮�����")

        line = line + 1
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line + 1,1,0) + "C5" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C4","������")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C6:R" + str(line + 1,1,0) + "C7" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,TRUE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C6","�६� ࠡ��� � �����")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C8:R" + str(line + 1,1,0) + "C9" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C8","������")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C10:R" + str(line + 1,1,0) + "C11" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C10","�६� �����")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C12:R" + str(line + 1,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C12","%")

        line = line + 1
        s = keyfirst(subgroups)
        while( strlen(trim(s)) != 0 )
          line = line + 1
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C3" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
          dde_execute("[FORMAT.FONT("+chr(34)+"Times New Roman"+chr(34)+",11,FALSE,FALSE,FALSE,FALSE,1)]")
          dde_poke("R" + str(line,1,0) + "C1",trim(s))

          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C5" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C4",subgroups[s][1])

          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C6:R" + str(line,1,0) + "C7" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C6",subgroups[s][2])

          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C8:R" + str(line,1,0) + "C9" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C8",subgroups[s][3])

          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C10:R" + str(line,1,0) + "C11" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C10"," ")

          dde_poke("R" + str(line,1,0) + "C12"," ")

          s = keynext(subgroups,s)
        wend

        dde_execute("[SELECT(" + chr(34) + "R" + str(start_line,1,0) + "C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[BORDER(1,1,1,1,1)]")

        line = line + 2
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C2:R" + str(line,1,0) + "C9" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C2","�㪮����⥫� ���ࠧ�������______________________________")

        dde_execute("[WORKBOOK.ACTIVATE(" + chr(34) + "����" + chr(34) + ")]")

        fio_maineng = fio_ini(trim('��.�������'.org))
        dde_poke("R2C1",trim('���������'.org))
        dde_poke("R2C6",trim('���������'.org))
        dde_poke("R5C1",fio_maineng)
        dde_poke("R5C6",fio_maineng)
        dde_poke("R7C6",chr(34) + "___" + chr(34) + " _____________ " + year + " �.")
        dde_poke("R7C1",chr(34) + "___" + chr(34) + " _____________ " + year + " �.")
        dde_poke("R7C6",chr(34) + "___" + chr(34) + " _____________ " + year + " �.")
        dde_poke("R13C1","� �த������ ࠡ�� �� " + name_months[month][1] + " " + year + " ����")
        kolsotr = 0
        chelchas = 0
        s = keyfirst(people)
        while( strlen(trim(s)) != 0 )
rem          print s,people[s]
          kolsotr = kolsotr + 1
          chelchas = chelchas + people[s]
          s = keynext(people,s)
        wend
        dde_poke("R18C1","����� ��ਮ� � " + date_b + " �. �� " + date_e + " �. " + str(kolsotr,1,0) + " 祫., " + str(chelchas,1,0) + " �/��.")

        dde_poke("R23C1",trim('��������'.org))
        fio_zmaineng = fio_ini(trim('���.��.���'.org))
        dde_poke("R24C5",fio_zmaineng)

        fio_rukotd = fio_ini(trim('��-���'.org))
        dde_poke("R27C5",fio_rukotd)
        dde_poke("R31C4"," ����� " + year + " �.")

        dde_execute("[SELECT(" + chr(34) + "R1C1" + chr(34) + ")]")
        dde_execute("[APP.RESTORE()]")
        dde_close

        end
