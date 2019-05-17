$include pathxls
        b = open("buffer")
        getfirst(b)
        date_start = trim('Строка_2'.b)
        date_finish = trim('Строка_3'.b)
        flag_two = 0
        if substr(date_start,1,4) != substr(date_finish,1,4) then
          flag_two = 1
        end if

        year = substr('Строка_3'.b,1,4)
        month = substr('Строка_3'.b,5,2)
        year_month = substr('Строка_3'.b,1,6)

        name_months = keyarray
        name_months["01"] = array(3)
        name_months["01"][1] = "Январь"
        name_months["01"][2] = "январе"
        name_months["01"][3] = 31

        name_months["02"] = array(3)
        name_months["02"][1] = "Февраль"
        name_months["02"][2] = "феврале"
        name_months["02"][3] = 28

        name_months["03"] = array(3)
        name_months["03"][1] = "Март"
        name_months["03"][2] = "марте"
        name_months["03"][3] = 31

        name_months["04"] = array(3)
        name_months["04"][1] = "Апрель"
        name_months["04"][2] = "апреле"
        name_months["04"][3] = 30

        name_months["05"] = array(3)
        name_months["05"][1] = "Май"
        name_months["05"][2] = "мае"
        name_months["05"][3] = 31

        name_months["06"] = array(3)
        name_months["06"][1] = "Июнь"
        name_months["06"][2] = "июне"
        name_months["06"][3] = 30

        name_months["07"] = array(3)
        name_months["07"][1] = "Июль"
        name_months["07"][2] = "июле"
        name_months["07"][3] = 31

        name_months["08"] = array(3)
        name_months["08"][1] = "Август"
        name_months["08"][2] = "августе"
        name_months["08"][3] = 31

        name_months["09"] = array(3)
        name_months["09"][1] = "Сентябрь"
        name_months["09"][2] = "сентябре"
        name_months["09"][3] = 30

        name_months["10"] = array(3)
        name_months["10"][1] = "Октябрь"
        name_months["10"][2] = "октябре"
        name_months["10"][3] = 31

        name_months["11"] = array(3)
        name_months["11"][1] = "Ноябрь"
        name_months["11"][2] = "ноябре"
        name_months["11"][3] = 30

        name_months["12"] = array(3)
        name_months["12"][1] = "Декабрь"
        name_months["12"][2] = "декабре"
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
          w_message("Ошибка!","Невозможно собрать данные за указанный месяц, так как он не активирован в настройках!",0)
          end
        end if

        w_open(300,250,270,160,"Сделайте выбор")
        w_button(6,80,15,100,25,"План")
        w_button(7,80,50,100,25,"Факт")
        w_button(10,80,85,100,25,"Отмена")

        po = open("plan_otchet")
        svod_plan = keyarray

function plan_fact
        result = 1
        local i
        a = regex_split("[:;]",'Значения'.po)
        for i = 1 to a[0]
          if strpos("R",a[i]) != 0 | strpos("C",a[i]) != 0
            if keyfind(svod_plan,a[i]) = 0
              svod_plan[a[i]] = a[i + 1]
              i = i + 1
            else
              w_message("Ошибка!","Одинаковые адреса у ячеек! " + a[i],0)
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
            button = w_message("Внимание!","Обнаружено, что ранее формировался отчет за данный месяц! Заменить данные по плану новыми?",4)
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
            w_message("Внимание!","Обнаружено, что отчет не формировался по плановым показателям! Будут использованы в плане фактические данные!",0)
          end if
        else
          w_close
          w_message("Внимание!","Выполнение отчета отменено!",0)
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
        dde_execute("[OPEN(" + chr(34) + pathxls + "План-отчет.xls" + chr(34) + ")]")
        dde_execute("[APP.MINIMIZE()]")

        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from OST_OBORUD where ""Дата_ввода"" IS NOT NULL AND (""ТО_мес_год"" IS NULL OR ""ТО_мес_год""='0')" )
        kolvo = sz["Kolvo"]
        close(sz)
rem     лист ТО

        stanki_to = keyarray

        if kolvo > 0
          ns = sql_query(Upper(getpath),"SELECT * from OST_OBORUD where ""Дата_ввода"" IS NOT NULL AND (""ТО_мес_год"" IS NULL OR ""ТО_мес_год""='0') ORDER BY ""Ном_объект"" ASC, ""Артикул"" ASC")
          dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "ТО" + chr(34) + ")]")
          dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
          dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",9,FALSE,FALSE,FALSE,FALSE,1)]")
          dde_execute("[CLEAR()]")
          dde_execute("[BORDER(0,0,0,0,0)]")
          dde_poke("R2C1","отдела №28, согласно годового графика ТО в " + name_months[month][2] + " месяце " + year + " г.")

          count_wp = 0

          w_progress("Лист ТО...",0,kolvo)
          getfirst(ns)
          while( eof(ns) = 0 )
            count_wp = count_wp + 1
            w_progress(count_wp)
            id_oborud = ns["ID_оборуд"]
rem            print "new stanki ",ns["Инв_номер"]
            rm = sql_query(Upper(getpath),"SELECT * from WORK_OBORUD where ""ID_оборуд""='" + str(id_oborud,1,0) + "' AND SUBSTRING(""Дата_н_пл"" FROM 1 FOR 4)='" + year + "' ORDER BY ""ID_раб_об"" ASC")
            getfirst(rm)
            while( eof(rm) = 0 )
              if substr(rm["Дата_н_пл"],1,6) = year_month then
rem                print rm["Дата_н_пл"],rm["ID_оборуд"],rm["ID_раб_об"],rm["ID_раб_уз"]
                if strlen(trim(rm["Этап_ТО"])) != 0 then
                  if strlen(trim(ns["Ном_объект"])) != 0
                    if keyfind(stanki_to,ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]) = 0
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]] = array(20)
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][1] = ns["Ном_объект"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][2] = ns["Инв_номер"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][3] = ns["Оборуд"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][4] = ns["Артикул"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][5] = rm["Этап_ТО"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][6] = rm["Трудоем_пл"] / 2.5
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][7] = rm["Трудоем_пл"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][8] = rm["Дата_н_ф"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][9] = rm["Дата_к_ф"]
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][10] = 0
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][11] = 0

                      if rm["Дата_н_ф"] < date_b & strlen(trim(rm["Дата_к_ф"])) = 0 | rm["Дата_к_ф"] > date_e
                      else
                        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_раб_уз""='" + str(rm["ID_раб_уз"],1,0) + "'")
                        kolvo_wu = sz["Kolvo"]
                        close(sz)
                        if kolvo_wu > 0 then
                          wu = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_раб_уз""='" + str(rm["ID_раб_уз"],1,0) + "'")
                          getfirst(wu)
                          stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][10] = wu["Длит_факт"]
                          stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][11] = wu["Трудоем_ф"]
                          close(wu)
                        end if
                      end if
                    else
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][6] = stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][6] + rm["Трудоем_пл"] / 2.5
                      stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][7] = stanki_to[ns["Ном_объект"] + ns["Артикул"] + ns["Инв_номер"]][7] + rm["Трудоем_пл"]
                    end if
                  else
                    print "У станка с инвентарным номером " + ns["Инв_номер"] + " не указан цех. Станок не включен в отчет!"
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
            dde_poke("R" + str(line,1,0) + "C1","Выполнение работ в рамках ТРМ")
            dde_poke("R" + str(line,1,0) + "C8",100)

            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C6" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","Итого")
            dde_poke("R" + str(line,1,0) + "C7",itogi[7])
            dde_poke("R" + str(line,1,0) + "C8",itogi[8] + 100)
            dde_poke("R" + str(line,1,0) + "C11",itogi[11])
            dde_poke("R" + str(line,1,0) + "C12",itogi[12])

            dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
            dde_execute("[BORDER(1,1,1,1,1)]")

            line = line + 2
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C4" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","ТО-0 ежемесячное обслуживание")

            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C4" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","ТО-1 полугодовое обслуживание")

            line = line + 1
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C4" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(6,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C1","ТО-2 ежегодное обслуживание")

            line = line + 2
            dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C7" + chr(34) + ")]")
            dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
            dde_poke("R" + str(line,1,0) + "C4","Руководитель подразделения______________________________")
          end if
          dde_execute("[SELECT(" + chr(34) + "R1C1:R1C1" + chr(34) + ")]")
        else
          w_message("Ошибка!","Не найдены станки, которым требуется выполнить ТО в текущем месяце!",0)
        end if

rem        print "SELECT * from OST_OBORUD where ""Дата_ввода"" IS NOT NULL AND (""ТО_мес_год"" IS NULL OR ""ТО_мес_год""='1') ORDER BY ""ID_оборуд"" ASC"
        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from OST_OBORUD where ""Дата_ввода"" IS NOT NULL AND ""ТО_мес_год""='1' AND ""ВклВОтчет""='1'")
        kolvo = sz["Kolvo"]
        close(sz)

rem     Листы ТР и Осмотр
        if kolvo > 0 then
          os = sql_query(Upper(getpath),"SELECT * from OST_OBORUD where ""Дата_ввода"" IS NOT NULL AND ""ТО_мес_год""='1' AND ""ВклВОтчет""='1' ORDER BY ""Ном_объект"" ASC, ""Артикул"" ASC")
          tecrem = keyarray
          osmotr = keyarray

          count_wp = 0
          w_progress("Листы Осмотр и ТР...",0,kolvo)

          getfirst(os)
          while( eof(os) = 0 )
            count_wp = count_wp + 1
            w_progress(count_wp)
rem            print os["Инв_номер"]," // ",os["Оборуд"]," // ",os["Дата_ввода"]," // ",os["ВклВОтчет"]," // "

            id_oborud = os["ID_оборуд"]
            sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_OBORUD where ""ID_оборуд""='" + str(id_oborud,1,0) + "' AND SUBSTRING(""Дата_н_пл"" FROM 1 FOR 4)='" + year + "'")
            kolvo_wo = sz["Kolvo"]
            close(sz)

            flag_osm = 0
            flag_rem = 0
            if kolvo_wo > 0
              rm = sql_query(Upper(getpath),"SELECT * from WORK_OBORUD where ""ID_оборуд""='" + str(id_oborud,1,0) + "' AND SUBSTRING(""Дата_н_пл"" FROM 1 FOR 4)='" + year + "' ORDER BY ""ID_раб_об"" ASC")
              getfirst(rm)
              while( eof(rm) = 0 & flag_osm = 0 & flag_rem = 0 )
                if substr(rm["Дата_н_пл"],1,6) = year_month then
rem                  print rm["Дата_н_пл"],rm["ID_оборуд"],rm["ID_раб_об"],rm["ID_раб_уз"]
                  if strlen(trim(rm["Этап_ТО"])) != 0 then
                    if strlen(trim(os["Ном_объект"])) != 0
                      if find(vid,1,1,os["ID_марка"]) = 0
                        if strpos("0",rm["Этап_ТО"]) != 0
                          if keyfind(osmotr,str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]) = 0 then
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]] = array(10)
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][1] = os["Ном_объект"]
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][2] = os["Инв_номер"]
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][3] = os["Оборуд"]
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][4] = os["Артикул"]
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][5] = 'КРСм'.vid * 0.4
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][6] = 'КРСм'.vid * 0.85 + 'КРСэ'.vid * 0.2
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][7] = rm["Дата_н_ф"]
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][8] = rm["Дата_к_ф"]
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][9] = 0
                            osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][10] = 0

                            sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_раб_уз""='" + str(rm["ID_раб_уз"],1,0) + "'")
                            kolvo_wu = sz["Kolvo"]
                            close(sz)
                            if kolvo_wu > 0 then
                              wu = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_раб_уз""='" + str(rm["ID_раб_уз"],1,0) + "'")
                              getfirst(wu)
                              osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][9] = wu["Длит_факт"]
                              osmotr[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][10] = wu["Трудоем_ф"]
                              close(wu)
                            end if
                          end if
                        else if strpos("1",rm["Этап_ТО"]) != 0
                          if keyfind(tecrem,str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]) = 0 then
rem                            print "TP = " + os["Инв_номер"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]] = array(10)
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][1] = os["Ном_объект"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][2] = os["Инв_номер"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][3] = os["Оборуд"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][4] = os["Артикул"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][5] = 'КРСм'.vid * 2
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][6] = 'КРСм'.vid * 6 + 'КРСэ'.vid * 1.5
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][7] = rm["Дата_н_ф"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][8] = rm["Дата_к_ф"]
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][9] = 0
                            tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][10] = 0

                            sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_раб_уз""='" + str(rm["ID_раб_уз"],1,0) + "'")
                            kolvo_wu = sz["Kolvo"]
                            close(sz)
                            if kolvo_wu > 0 then
                              wu = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_раб_уз""='" + str(rm["ID_раб_уз"],1,0) + "'")
                              getfirst(wu)
                              tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][9] = wu["Длит_факт"]
                              tecrem[str(rm["ID_раб_уз"],1,0) + os["Инв_номер"]][10] = wu["Трудоем_ф"]
                              close(wu)
                            end if
                          end if
                        else
                          print "Неизвестное ТО ",rm["Этап_ТО"],os["Инв_номер"]
                        end if
                      else
                        print "Ошибка! Не найдена ID_марка ",os["ID_марка"]
                      end if
                    else
                      print "У станка с инвентарным номером " + os["Инв_номер"] + " не указан цех. Станок не включен в отчет!"
                    end if
                  end if
                end if
                getnext(rm)
              wend
            else
              print "Не найдены данные в WORK_OBORUD ",id_oborud
            end if
            close(rm)
            getnext(os)
          wend
          close(os)
          w_progress
        end if

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "Осмотр" + chr(34) + ")]")
        dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",9,FALSE,FALSE,FALSE,FALSE,1)]")
        dde_execute("[CLEAR()]")
        dde_execute("[BORDER(0,0,0,0,0)]")
        dde_poke("R2C1","отдела №28, согласно годового графика ППР в " + name_months[month][2] + " месяце " + year + " г.")
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
          dde_poke("R" + str(line,1,0) + "C6","О")
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
          dde_poke("R" + str(line,1,0) + "C1","Итого")
          dde_poke("R" + str(line,1,0) + "C7",itogi[1])
          dde_poke("R" + str(line,1,0) + "C8",itogi[2])
          dde_poke("R" + str(line,1,0) + "C11",itogi[3])
          dde_poke("R" + str(line,1,0) + "C12",itogi[4])

          dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
          dde_execute("[BORDER(1,1,1,1,1)]")

          line = line + 3
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C6" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C4","Руководитель подразделения______________________________")
        end if
        dde_execute("[SELECT(" + chr(34) + "R1C1:R1C1" + chr(34) + ")]")

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "ТР" + chr(34) + ")]")
        dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",9,FALSE,FALSE,FALSE,FALSE,1)]")
        dde_execute("[CLEAR()]")
        dde_execute("[BORDER(0,0,0,0,0)]")
        dde_poke("R2C1",name_months[month][2] + " месяце " + year + " г. Отдел №28")
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
          dde_poke("R" + str(line,1,0) + "C6","ТР")
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
          dde_poke("R" + str(line,1,0) + "C1","Итого")
          dde_poke("R" + str(line,1,0) + "C7",itogi[1])
          dde_poke("R" + str(line,1,0) + "C8",itogi[2])
          dde_poke("R" + str(line,1,0) + "C11",itogi[3])
          dde_poke("R" + str(line,1,0) + "C12",itogi[4])

          dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
          dde_execute("[BORDER(1,1,1,1,1)]")

          line = line + 3
          dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C7" + chr(34) + ")]")
          dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
          dde_poke("R" + str(line,1,0) + "C4","Руководитель подразделения______________________________")
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

        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from OST_OBORUD where ""Дата_ввода"" IS NOT NULL")
        kolvo = sz["Kolvo"]
        close(sz)
        if kolvo > 0 then
          all = sql_query(Upper(getpath),"SELECT * from OST_OBORUD where ""Дата_ввода"" IS NOT NULL")
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
            if keyfind(stops,all["Ном_объект"]) = 0
              stops[all["Ном_объект"]] = array(40)
              for i = 1 to 40
                stops[all["Ном_объект"]][i] = 0
              next
              stops[all["Ном_объект"]][23] = stops[all["Ном_объект"]][23] + 1
            else
              stops[all["Ном_объект"]][23] = stops[all["Ном_объект"]][23] + 1
            end if
            if find(vid,1,1,all["ID_марка"]) = 0
              if strlen(trim('Подгруппа'.vid)) != 0
                flag_sg = 1
                if keyfind(subgroups,trim('Подгруппа'.vid)) = 0
                  subgroups[trim('Подгруппа'.vid)] = array(5)
                  subgroups[trim('Подгруппа'.vid)][1] = 1
                  subgroups[trim('Подгруппа'.vid)][2] = 0
                  subgroups[trim('Подгруппа'.vid)][3] = 0
                else
                  subgroups[trim('Подгруппа'.vid)][1] = subgroups[trim('Подгруппа'.vid)][1] + 1
                end if
              else
                print "34 // ",all["Инв_номер"]," отсутствует название подгруппы"
              end if
            else
              print "33"
            end if

            if flag_two = 0
              if findge(yo1,2,1,all["ID_оборуд"],date_start) = 0 then
                while( eof(yo1) = 0 & all["ID_оборуд"] = 'ID_оборуд'.yo1 & 'Дата'.yo1 <= date_finish )
                  stops[all["Ном_объект"]][24] = stops[all["Ном_объект"]][24] + 'Нараб_план'.yo1
                  if flag_sg = 1 then
                    subgroups[trim('Подгруппа'.vid)][2] = subgroups[trim('Подгруппа'.vid)][2] + 'Нараб_план'.yo1
                  end if
                  getnext(yo1)
                wend
              end if
            else
              if findge(yo1,2,1,all["ID_оборуд"],date_start) = 0 then
                while( eof(yo1) = 0 & all["ID_оборуд"] = 'ID_оборуд'.yo1 & 'Дата'.yo1 <= date_finish )
                  stops[all["Ном_объект"]][24] = stops[all["Ном_объект"]][24] + 'Нараб_план'.yo1
                  if flag_sg = 1 then
                    subgroups[trim('Подгруппа'.vid)][2] = subgroups[trim('Подгруппа'.vid)][2] + 'Нараб_план'.yo1
                  end if
                  getnext(yo1)
                wend
              end if
              if findge(yo2,2,1,all["ID_оборуд"],date_start) = 0 then
                while( eof(yo2) = 0 & all["ID_оборуд"] = 'ID_оборуд'.yo2 & 'Дата'.yo2 <= date_finish )
                  stops[all["Ном_объект"]][24] = stops[all["Ном_объект"]][24] + 'Нараб_план'.yo2
                  if flag_sg = 1 then
                    subgroups[trim('Подгруппа'.vid)][2] = subgroups[trim('Подгруппа'.vid)][2] + 'Нараб_план'.yo2
                  end if
                  getnext(yo2)
                wend
              end if
            end if
            getnext(all)
          wend
          close(all)
        end if

rem     Табель рабочего времени
        common_fond = 0
        otpuska = 0
        blist = 0
        proch_lost = 0
        tmr = 0
        kolsotr = 0
        als = open("all_sotr")
        grvs = open("grafic_rv_sotrud")
        typetab = open("type_tabel")

rem        print "SELECT * from TABELRV where ""Дата"">='" + date_b + "' AND ""Дата""<='" + date_e + "' ORDER BY ""ID_сотр,Дата"" ASC"

        sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from TABELRV where ""Дата"">='" + date_b + "' AND ""Дата""<='" + date_e + "'")
        kolvo = sz["Kolvo"]
        close(sz)
        if kolvo > 0 then
          trv = sql_query(Upper(getpath),"SELECT * from TABELRV where ""Дата"">='" + date_b + "' AND ""Дата""<='" + date_e + "' ORDER BY ""ID_сотр"" ASC, ""Дата"" ASC")
          getfirst(trv)
          while( eof(trv) = 0 )
            if find(als,1,1,'ID_сотр'.trv) = 0
              if 'НеВклПО'.als = 0
                print 'Сотрудник'.als,"\\",'Дата'.trv,"\\",'ГрРВ_Сот'.trv
                if keyfind(people,str('ID_сотр'.als,1,0) + 'Сотрудник'.als) = 0 then
                  people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] = 0
                end if
                if strpos("ВЫХОДНОЙ",Upper('ГрРВ_Сот'.trv)) != 0
                else
                  if find(grvs,1,1,'ID_ГрРВСот'.trv) = 0
                    fl_smena = 0
                    for i = 1 to 6
                      field_start = "Нач_" + str(i,1,0)
                      field_finish = "Кон_" + str(i,1,0)
                      if strlen(trim('field_start'.grvs)) != 0 & strlen(trim('field_finish'.grvs)) != 0 then
                        if fl_smena = 0 then
                          fl_smena = 1
                        end if
                      end if
                    next
                    if fl_smena = 1 then
                      for i = 1 to 6
                        field_start = "Нач_" + str(i,1,0)
                        field_finish = "Кон_" + str(i,1,0)
                        if strlen(trim('field_start'.trv)) != 0 & strlen(trim('field_finish'.trv)) != 0 then
                          common_fond = common_fond + kol_chasov('field_start'.trv,'field_finish'.trv)
                          people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] = people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] + kol_chasov('field_start'.trv,'field_finish'.trv)
                          if strpos("28-6",'Бригада'.als) != 0 then
                            pf[1][4] = pf[1][4] + kol_chasov('field_start'.trv,'field_finish'.trv)
                          end if
                        end if
                      next
                    end if
                  else if find(typetab,1,1,'ID_ГрРВСот'.trv) = 0
                    if 'Отпуск'.typetab = 1
  rem                    print "Отпуск ",otpuska
                      if 'День_цикл'.trv = 1 | 'День_цикл'.trv = 2
                        otpuska = otpuska + 11
                        people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] = people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] + 11
                      else if 'День_цикл'.trv = 0 & trim('Ден_нед'.trv) != "суббота" & trim('Ден_нед'.trv) != "воскресенье"
                        otpuska = otpuska + 8
                        people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] = people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] + 8
                      end if
rem                      print "Отпуск ",otpuska
                    else if strpos("БОЛЬНИЧНЫЙ",Upper('Знач_таб'.typetab)) != 0
                      if 'День_цикл'.trv = 1 | 'День_цикл'.trv = 2
                        blist = blist + 11
                        people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] = people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] + 11
                      else if 'День_цикл'.trv = 0 & trim('Ден_нед'.trv) != "суббота" & trim('Ден_нед'.trv) != "воскресенье"
                        blist = blist + 8
                        people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] = people[str('ID_сотр'.als,1,0) + 'Сотрудник'.als] + 8
                      end if
                    else
                      proch_lost = proch_lost + 8
                    end if
                  end if
                end if
              else
              end if
            else
              w_message("Ошибка!","Не найден сотрудник в отделе кадров " + trim('Сотрудник'.trv) + "!",0)
            end if
            getnext(trv)
          wend
          close(trv)
        end if

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "ТМР" + chr(34) + ")]")
        dde_execute("[SELECT(" + chr(34) + "R8C1:R1000C12" + chr(34) + ")]")
        dde_execute("[FORMAT.FONT(" + chr(34) + "Times New Roman" + chr(34) + ",12,FALSE,FALSE,FALSE,FALSE,1)]")
        dde_execute("[CLEAR()]")
        dde_execute("[BORDER(0,0,0,0,0)]")

        dde_poke("R3C1","отдела № 28, в " + name_months[month][2] + " месяце " + year + " г.")

        line = 8
        dde_poke("R8C6","Демонтаж и перестановка старого оборудования")
        dde_poke("R8C8",num(str(pf[1][4] / 2,1,0)))
        dde_poke("R9C6","Приемка нового оборудования")
        dde_poke("R9C8",pf[1][4] - num(str(pf[1][4] / 2,1,0)))
        line = line + 3
        dde_poke("R" + str(line,1,0) + "C8",pf[1][4])
        dde_execute("[SELECT(" + chr(34) + "R8C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[BORDER(1,1,1,1,1)]")

        line = line + 2
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C6" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C4","Руководитель подразделения______________________________")

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "ВР" + chr(34) + ")]")
        dde_poke("R3C1","ремонт оборудования в " + name_months[month][2] + " месяце " + year + " г.")
        line = 8
rem        print "SELECT * from WORK_UNIT where ( ""ID_авария"" IS NOT NULL OR ""ID_авария"">'0' ) AND ( ( ""Дата_н_ф"">='" + date_b + "' AND ""Дата_н_ф""<='" + date_e + "' AND ""Дата_к_ф"">='" + date_b + "' AND ""Дата_к_ф""<='" + date_e + "' ) OR ( ""Дата_н_ф"">='" + date_b + "' AND ""Дата_н_ф""<='" + date_e + "' AND ( ""Дата_к_ф"">'" + date_e + "' OR ""Дата_к_ф"" IS NULL ) ) OR ( ""Дата_н_ф""<='" + date_b + "' AND ""Дата_к_ф"">='" + date_b + "' AND ""Дата_к_ф""<='" + date_e + "') OR ( ""Дата_н_ф""<'" + date_b + "' AND ( ""Дата_к_ф"" IS NULL OR ""Дата_к_ф"">'" + date_e + "') ) ) ORDER BY ""Дата_н_ф"" ASC"

        vr = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_UNIT where ""ID_авария"">'0' AND ( ( ""Дата_н_ф"">='" + date_b + "' AND ""Дата_н_ф""<='" + date_e + "' AND ""Дата_к_ф"">='" + date_b + "' AND ""Дата_к_ф""<='" + date_e + "' ) OR ( ""Дата_н_ф"">='" + date_b + "' AND ""Дата_н_ф""<='" + date_e + "' AND ( ""Дата_к_ф"">'" + date_e + "' OR ""Дата_к_ф"" IS NULL ) ) OR ( ""Дата_н_ф""<='" + date_b + "' AND ""Дата_к_ф"">='" + date_b + "' AND ""Дата_к_ф""<='" + date_e + "') OR ( ""Дата_н_ф""<'" + date_b + "' AND ( ""Дата_к_ф"" IS NULL OR ""Дата_к_ф"">'" + date_e + "') ) ) ")
        kolvo = vr["Kolvo"]
        close(vr)

        if kolvo > 0 then
          count = 0

          for i = 1 to 30
            itogi[i] = 0
          next
          ost = open("ost_oborud")
          vr = sql_query(Upper(getpath),"SELECT * from WORK_UNIT where ""ID_авария"">'0' AND ( ( ""Дата_н_ф"">='" + date_b + "' AND ""Дата_н_ф""<='" + date_e + "' AND ""Дата_к_ф"">='" + date_b + "' AND ""Дата_к_ф""<='" + date_e + "' ) OR ( ""Дата_н_ф"">='" + date_b + "' AND ""Дата_н_ф""<='" + date_e + "' AND ( ""Дата_к_ф"">'" + date_e + "' OR ""Дата_к_ф"" IS NULL ) ) OR ( ""Дата_н_ф""<='" + date_b + "' AND ""Дата_к_ф"">='" + date_b + "' AND ""Дата_к_ф""<='" + date_e + "') OR ( ""Дата_н_ф""<'" + date_b + "' AND ( ""Дата_к_ф"" IS NULL OR ""Дата_к_ф"">'" + date_e + "') ) ) ORDER BY ""Дата_н_ф"" ASC")
          getfirst(vr)
          while( eof(vr) = 0 )
rem            print vr["ID_раб_уз"]
            if findge(ost,1,3,vr["ID_узла"]) = 0 then
              while( eof(ost) = 0 & vr["ID_узла"] = 'ID_узла'.ost )
                line = line + 1
                count = count + 1
                dde_poke("R" + str(line,1,0) + "C1","'" + str(count,1,0))
                dde_poke("R" + str(line,1,0) + "C2","'" + trim('Ном_объект'.ost))
                dde_poke("R" + str(line,1,0) + "C3","'" + trim('Инв_номер'.ost))
rem                print trim('Инв_номер'.ost)
                dde_poke("R" + str(line,1,0) + "C4",trim('Оборуд'.ost))
                dde_poke("R" + str(line,1,0) + "C5","'" + trim('Артикул'.ost))
rem                dde_poke("R" + str(line,1,0) + "C6",trim(vr["Прим_р_уз"]))

                sz = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from WORK_OBORUD where ""ID_раб_уз""='" + str(vr["ID_раб_уз"],1,0) + "'")
                kolvo_wo = sz["Kolvo"]
                close(sz)
                if kolvo_wo > 0 then
                  wo = sql_query(Upper(getpath),"SELECT * from WORK_OBORUD where ""ID_раб_уз""='" + str(vr["ID_раб_уз"],1,0) + "' ORDER BY ""ID_раб_об"" ASC")
                  list_work = keyarray
                  getfirst(wo)
                  while( eof(wo) = 0 )
                    if keyfind(list_work,str(wo["ID_раб_об"],1,0)) = 0 then
                      list_work[str(wo["ID_раб_об"],1,0)] = array(3)
                      list_work[str(wo["ID_раб_об"],1,0)][1] = trim(wo["Прим_р_о"])
                      list_work[str(wo["ID_раб_об"],1,0)][2] = trim(wo["Вид_работ"])
                      list_work[str(wo["ID_раб_об"],1,0)][3] = trim(wo["ОписТехнол"])
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
rem                print vr["Дата_н_ф"]," // ",date_b," // ",strlen(trim(vr["Дата_к_ф"]))," // ",vr["Дата_к_ф"]," // ",date_e,"//"
                if ( vr["Дата_н_ф"] < substr(date_b,7,4) + substr(date_b,4,2) + substr(date_b,1,2) & strlen(trim(vr["Дата_к_ф"])) = 0 ) | ( vr["Дата_к_ф"] > substr(date_e,7,4) + substr(date_e,4,2) + substr(date_e,1,2) )
                else
rem print "11"
                  dde_poke("R" + str(line,1,0) + "C7",vr["Длит_план"])
                  dde_poke("R" + str(line,1,0) + "C8",vr["Трудоем_пл"])
                  if substr(date_b,7,4) = substr(vr["Дата_к_ф"],1,4) | strlen(trim(vr["Дата_к_ф"])) = 0
                    tablname = "Y" + substr(date_b,7,4) + "_OBORUD"
rem                    print tablname,"// ","SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4) + "'"
                    if strlen(trim(vr["Дата_к_ф"])) = 0
                      kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='" + date_e + "'")
                    else
                      kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4) + "'")
                    end if
                    kolyobor = kol_yo["Kolvo"]
                    close(kol_yo)

                    if kolyobor > 0 then
                      if strlen(trim(vr["Дата_к_ф"])) = 0
                        yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='" + date_e + "'")
                      else
                        yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4) + "'")
                      end if
                      prostoi = 0
                      getfirst(yobor)
                      while( eof(yobor) = 0 )
                        if vr["Дата_к_ф"] != yobor["Дата"]
                          prostoi = prostoi + yobor["Раб_план"]
                        else
                          if vr["Дата_к_ф"] = vr["Дата_н_ф"]
                            prostoi = prostoi + ( num(substr(vr["Время_к_ф"],1,2)) * 60 + num(substr(vr["Время_к_ф"],4,2)) - ( num(substr(vr["Время_н_ф"],1,2)) * 60 + num(substr(vr["Время_н_ф"],4,2)) ) ) / 60
                          else
                            prostoi = prostoi + (( num(substr(vr["Время_к_ф"],1,2)) - 8 ) * 60 + num(substr(vr["Время_к_ф"],4,2))) / 60
                          end if
                        end if
                        getnext(yobor)
                      wend
                      close(yobor)
                    end if
                  else
                    prostoi = 0
                    tablname = "Y" + substr(date_b,7,4) + "_OBORUD"
rem                    print tablname,"// ","SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='31.12." + substr(date_b,7,4) + "'"
                    kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='31.12." + substr(date_b,7,4) + "'")
                    kolyobor = kol_yo["Kolvo"]
                    close(kol_yo)

                    if kolyobor > 0 then
                      yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='" + date_b + "' AND ""Дата""<='31.12." + substr(date_b,7,4) + "'")
                      getfirst(yobor)
                      while( eof(yobor) = 0 )
                        if vr["Дата_к_ф"] != yobor["Дата"]
                          prostoi = prostoi + yobor["Раб_план"]
                        else
                          if vr["Дата_к_ф"] = vr["Дата_н_ф"]
                            prostoi = prostoi + ( num(substr(vr["Время_к_ф"],1,2)) * 60 + num(substr(vr["Время_к_ф"],4,2)) - ( num(substr(vr["Время_н_ф"],1,2)) * 60 + num(substr(vr["Время_н_ф"],4,2)) ) ) / 60
                          else
                            prostoi = prostoi + (( num(substr(vr["Время_к_ф"],1,2)) - 8 ) * 60 + num(substr(vr["Время_к_ф"],4,2))) / 60
                          end if
                        end if
                        getnext(yobor)
                      wend
                      close(yobor)
                    end if

                    tablname = "Y" + substr(vr["Дата_к_ф"],1,4) + "_OBORUD"
rem                    print tablname,"// ","SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='01.01." + substr(vr["Дата_к_ф"],1,4) + "' AND ""Дата""<='" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4)
                    kol_yo = sql_query(Upper(getpath),"SELECT count(*) as Kolvo from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='01.01." + substr(vr["Дата_к_ф"],1,4) + "' AND ""Дата""<='" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4))
                    kolyobor = kol_yo["Kolvo"]
                    close(kol_yo)

                    if kolyobor > 0 then
                      yobor = sql_query(Upper(getpath),"SELECT * from " + tablname + " where ""ID_оборуд""='" + str('ID_оборуд'.ost,1,0) + "' AND ""Дата"">='01.01." + substr(vr["Дата_к_ф"],1,4) + "' AND ""Дата""<='" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4))
                      getfirst(yobor)
                      while( eof(yobor) = 0 )
                        if vr["Дата_к_ф"] != yobor["Дата"]
                          prostoi = prostoi + yobor["Раб_план"]
                        else
                          if vr["Дата_к_ф"] = vr["Дата_н_ф"]
                            prostoi = prostoi + ( num(substr(vr["Время_к_ф"],1,2)) * 60 + num(substr(vr["Время_к_ф"],4,2)) - ( num(substr(vr["Время_н_ф"],1,2)) * 60 + num(substr(vr["Время_н_ф"],4,2)) ) ) / 60
                          else
                            prostoi = prostoi + (( num(substr(vr["Время_к_ф"],1,2)) - 8 ) * 60 + num(substr(vr["Время_к_ф"],4,2))) / 60
                          end if
                        end if
                        getnext(yobor)
                      wend
                      close(yobor)
                    end if
                  end if
                  dde_poke("R" + str(line,1,0) + "C11",prostoi)

rem                  dde_poke("R" + str(line,1,0) + "C11",vr["Длит_факт"])
                  dde_poke("R" + str(line,1,0) + "C12",vr["Трудоем_ф"])
                  itogi[1] = itogi[1] + vr["Длит_план"]
                  itogi[2] = itogi[2] + vr["Трудоем_пл"]
                  itogi[3] = itogi[3] + vr["Длит_факт"]
                  itogi[4] = itogi[4] + vr["Трудоем_ф"]
                end if

                if strlen(trim(vr["Дата_н_ф"])) != 0 then
                  dde_poke("R" + str(line,1,0) + "C9","'" + substr(vr["Дата_н_ф"],7,2) + "." + substr(vr["Дата_н_ф"],5,2) + "." + substr(vr["Дата_н_ф"],1,4))
                end if
                if strlen(trim(vr["Дата_к_ф"])) != 0 then
                  dde_poke("R" + str(line,1,0) + "C10","'" + substr(vr["Дата_к_ф"],7,2) + "." + substr(vr["Дата_к_ф"],5,2) + "." + substr(vr["Дата_к_ф"],1,4))
                end if
                if find(vid,1,1,'ID_марка'.ost) = 0 then
                  if strlen(trim('Подгруппа'.vid)) != 0
                    if keyfind(subgroups,trim('Подгруппа'.vid)) = 0
                      subgroups[trim('Подгруппа'.vid)] = array(5)
                      subgroups[trim('Подгруппа'.vid)][1] = 1
                      subgroups[trim('Подгруппа'.vid)][2] = 0
                      subgroups[trim('Подгруппа'.vid)][3] = 1
                    else
                      subgroups[trim('Подгруппа'.vid)][3] = subgroups[trim('Подгруппа'.vid)][3] + 1
                    end if
                  else 
                    print "34-1 // ",'Инв_номер'.ost," отсутствует название подгруппы"
                  end if
                end if  
                getnext(ost)
              wend
            end if
            getnext(vr)
          wend
          close(vr)
          line = line + 20
          dde_poke("R" + str(line,1,0) + "C3","ИТОГО")
          dde_poke("R" + str(line,1,0) + "C7",itogi[1])
          dde_poke("R" + str(line,1,0) + "C8",itogi[2])
          dde_poke("R" + str(line,1,0) + "C11",itogi[3])
          dde_poke("R" + str(line,1,0) + "C12",itogi[4])

          dde_execute("[SELECT(" + chr(34) + "R9C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
          dde_execute("[BORDER(1,1,1,1,1)]")
        end if

        dde_execute("[WORKBOOK.ACTIVATE(" +chr(34) + "Сводный" + chr(34) + ")]")
        dde_poke("R2C1","учета выполнения работ за " + name_months[month][1] + " " + year + " г., ч/час.")
        dde_poke("R9C1","Простой технологического оборудования, час.")
rem  очистку сделать

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
        dde_poke("R" + str(line,1,0) + "C1","Итого")
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
              'Год'.po = year
              'Месяц'.po = month
              'Значения'.po = stroka
              insert(po)
            else
              'Год'.po = year
              'Месяц'.po = month
              'Значения'.po = stroka
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
        dde_poke("R" + str(line,1,0) + "C1","Допустимый расчетный")
        dde_poke("R" + str(line,1,0) + "C12",1333.0)
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[BORDER(1,1,1,1,1)]")

        line = line + 3
        start_line = line
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C1","Внеплановый простой оборудования по группам")

        line = line + 1
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line,1,0) + "C7" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C4","Общее количество")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C8:R" + str(line,1,0) + "C12" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C8","Количество единиц в простое")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C1:R" + str(line + 2,1,0) + "C3" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C1","Группа оборудования")

        line = line + 1
        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C4:R" + str(line + 1,1,0) + "C5" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C4","единиц")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C6:R" + str(line + 1,1,0) + "C7" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,TRUE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C6","время работы в месяц")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C8:R" + str(line + 1,1,0) + "C9" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C8","единиц")

        dde_execute("[SELECT(" + chr(34) + "R" + str(line,1,0) + "C10:R" + str(line + 1,1,0) + "C11" + chr(34) + ")]")
        dde_execute("[ALIGNMENT(7,FALSE,1,0)]")
        dde_poke("R" + str(line,1,0) + "C10","время простоя")

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
        dde_poke("R" + str(line,1,0) + "C2","Руководитель подразделения______________________________")

        dde_execute("[WORKBOOK.ACTIVATE(" + chr(34) + "Титул" + chr(34) + ")]")

        fio_maineng = fio_ini(trim('Гл.инженер'.org))
        dde_poke("R2C1",trim('ДолжГлИнж'.org))
        dde_poke("R2C6",trim('ДолжГлИнж'.org))
        dde_poke("R5C1",fio_maineng)
        dde_poke("R5C6",fio_maineng)
        dde_poke("R7C6",chr(34) + "___" + chr(34) + " _____________ " + year + " г.")
        dde_poke("R7C1",chr(34) + "___" + chr(34) + " _____________ " + year + " г.")
        dde_poke("R7C6",chr(34) + "___" + chr(34) + " _____________ " + year + " г.")
        dde_poke("R13C1","о проделаной работе за " + name_months[month][1] + " " + year + " года")
        kolsotr = 0
        chelchas = 0
        s = keyfirst(people)
        while( strlen(trim(s)) != 0 )
rem          print s,people[s]
          kolsotr = kolsotr + 1
          chelchas = chelchas + people[s]
          s = keynext(people,s)
        wend
        dde_poke("R18C1","Отчетный период с " + date_b + " г. по " + date_e + " г. " + str(kolsotr,1,0) + " чел., " + str(chelchas,1,0) + " ч/час.")

        dde_poke("R23C1",trim('ДолжнЗГИ'.org))
        fio_zmaineng = fio_ini(trim('Зам.Гл.Инж'.org))
        dde_poke("R24C5",fio_zmaineng)

        fio_rukotd = fio_ini(trim('Рук-льОтд'.org))
        dde_poke("R27C5",fio_rukotd)
        dde_poke("R31C4"," Челябинск " + year + " г.")

        dde_execute("[SELECT(" + chr(34) + "R1C1" + chr(34) + ")]")
        dde_execute("[APP.RESTORE()]")
        dde_close

        end
