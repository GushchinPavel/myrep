$include pathxls
        system("del auto.txt")
        system("del contact.txt")
        token = ""
        id_contact = ""
        id_lead = ""

        b = open("buffer",-2)
        getfirst(b)
        nomer = trim('��ப�_4'.b)
        exp = open("export_crm")
        expkl = open("export_crm_kl")
        expopl = open("export_crm_opl")
        ok = open("klients",-2)

function parser_auto
        local i
        fromfile("auto.txt",1)
        while( fromeof = 0 )
          input text
          if strpos("success",text) != 0
            words = regex_split("[{}:," + chr(34) + "]",text)
            for i = 1 to words[0]
              if strlen(trim(words[i])) != 0 then
                token = words[i]
              end if
            next
          else
            w_message("�訡��!","���ਧ��� �� ᠩ� �� �믮�����! �믮������ �ணࠬ�� ��ࢠ��! ���஡�� ���ਧ������� �����!",0)
            end
          end if
        wend
        result = 1
end function

function autorization
        result = 0

        system(pathxls + "curl\curl -H " + chr(34) + "Content-Type: application/json" + chr(34) + " -X POST -d @auto.json direct.lptracker.ru/login >auto.txt")
        start_time = gettime
        ss = num(substr(start_time,strlen(start_time) - 4,2)) * 60 + num(substr(start_time,strlen(start_time) - 1,2))
        flag = 0
        while( flag = 0 )
          if file_exists("auto.txt") != 0
            flag = 1
          else if ss - num(substr(gettime,strlen(gettime) - 4,2)) * 60 + num(substr(gettime,strlen(gettime) - 1,2)) > 5
            flag = 5
          end if
        wend
        if flag = 5
          if file_exists("auto.txt") = 0
            w_message("�訡��!","� �祭�� 5 ᥪ㭤 �� �� ����祭 �⢥� �� ᠩ�! �믮������ �ணࠬ�� ��⠭������!",0)
            end
          else
            parser_auto
          end if
        else
          if flag = 1 then
            parser_auto
          end if
        end if

        if strlen(trim(token)) = 0 then
          w_message("�訡��!","�� 㤠���� ������� ����! �믮������ �ணࠬ�� ��ࢠ��!",0)
          end
        end if
end function

function parser_contact
        local i,j
        fromfile("contact.txt",1)
        while( fromeof = 0 )
          input text
          if strpos("success",text) != 0 & strpos("[]",text) = 0
            pairs = regex_split("[\]\[{}:,]",text)
            for i = 1 to pairs[0]
              if strlen(trim(pairs[i])) != 0 then
                if strpos("id",pairs[i]) != 0 then
                  id_contact = pairs[i + 1]
                  i = pairs[0]
                end if
              end if
            next
          else
rem            w_message("�訡��!","�� ������ ⠪�� ������ � ���� ������ �� ᠩ�! �믮������ �ணࠬ�� ��ࢠ��!",0)
            print "�訡��!","�� ������ ⠪�� ������ � ���� ������ �� ᠩ�!"
          end if
        wend

        result = 1
end function

function findcontact(phone,email)
        result = 0
        if strlen(trim(phone)) != 0 & strlen(trim(email)) = 0
          system("curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/search?project_id=66913^&phone=" + phone + " >contact.txt")
rem          print "curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/search?project_id=66913^&phone=" + phone + " >contact.txt"
        else if strlen(trim(phone)) = 0 & strlen(trim(email)) != 0
          system("curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/search?project_id=66913^&email=" + email + " >contact.txt")
rem          print "curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/search?project_id=66913^&email=" + email + " >contact.txt"
        else
          system("curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/search?project_id=66913^&phone=" + phone + "^&email=" + email + " >contact.txt")
rem          print "curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/search?project_id=66913^&phone=" + phone + "^&email=" + email + " >contact.txt"
        end if

        start_time = gettime
        ss = num(substr(start_time,strlen(start_time) - 4,2)) * 60 + num(substr(start_time,strlen(start_time) - 1,2))
        flag = 0
        while( flag = 0 )
          if file_exists("contact.txt") != 0
            flag = 1
          else if ss - num(substr(gettime,strlen(gettime) - 4,2)) * 60 + num(substr(gettime,strlen(gettime) - 1,2)) > 5
            flag = 5
          end if
        wend
        if flag = 5
          if file_exists("contact.txt") = 0
            w_message("�訡��!","� �祭�� 5 ᥪ㭤 �� �� ����祭 �⢥� �� ᠩ�! ����� ������ " + trim(klients[s][1]) + " �� ���� ��।��� � CRM!",0)
          else
            parser_contact
          end if
        else
          if flag = 1 then
            parser_contact
          end if
        end if
        if strlen(trim(id_contact)) = 0 then
          print "�訡��!","�� 㤠���� ������� �����䨪��� ���⠪�! ����� ������ " + trim(klients[s][1]) + " �� ���� ��।��� � CRM!"
        end if
end function

function parser_leads
        local i,j

        result = 1
        fromfile("leads.txt",1)
        while( fromeof = 0 )
          input text
          if strpos("success",text) != 0
            pairs = regex_split("[,]",text)
            for i = 1 to pairs[0]
              if strlen(trim(pairs[i])) != 0 then
                if strpos("result",pairs[i]) != 0 then
                  znach = regex_split("[" + chr(34) + ":]",pairs[i])
                  for j = 1 to znach[0]
                    if strlen(trim(znach[j])) != 0 then
                      id_lead = znach[j]
                      i = pairs[0]
                    end if
                  next
                end if
              end if
            next
          else
            print "�訡��!","�� ������ ��� � ���� ������ �� ᠩ�!"
          end if
        wend
end function

function findleads
        result = 0
        system("curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/" + id_contact + "/leads >leads.txt")
rem        print "curl\curl -H " + chr(34) + "token: " + token + chr(34) + " direct.lptracker.ru/contact/" + id_contact + "/leads >leads.txt"
        start_time = gettime
        ss = num(substr(start_time,strlen(start_time) - 4,2)) * 60 + num(substr(start_time,strlen(start_time) - 1,2))
        flag = 0
        while( flag = 0 )
          if file_exists("leads.txt") != 0
            flag = 1
          else if ss - num(substr(gettime,strlen(gettime) - 4,2)) * 60 + num(substr(gettime,strlen(gettime) - 1,2)) > 5
            flag = 5
          end if
        wend
        if flag = 5
          if file_exists("leads.txt") = 0
            w_message("�訡��!","� �祭�� 5 ᥪ㭤 �� �� ����祭 �⢥� �� ᠩ�! ����� ������ " + trim(klients[s][1]) + " �� ���� ��।��� � CRM!",0)
          else
            parser_leads
          end if
        else
          if flag = 1 then
            parser_leads
          end if
        end if
        if strlen(trim(id_lead)) = 0 then
          print "�訡��!","�� 㤠���� ������� �����䨪��� ����!"
        end if
end function

function send_sum(summa)
        result = 0
        redirect("summa.json",1)
        print "{"
        print chr(34) + "category" + chr(34) + ": " + chr(34) + "1" + chr(34) + ","
        print chr(34) + "purpose" + chr(34) + ": " + chr(34) + "1" + chr(34) + ","
        print chr(34) + "sum" + chr(34) + ":" + str(summa,1,2)
        print "}"
        redirect

        flag_summa = 0
        system("curl\curl -H " + chr(34) + "Content-Type: application/json" + chr(34) + " -H " + chr(34) + "token: " + token + chr(34) + " -X POST -d @summa.json direct.lptracker.ru/lead/" + id_lead + "/payment >pay.txt")
        start_time = gettime
        ss = num(substr(start_time,strlen(start_time) - 4,2)) * 60 + num(substr(start_time,strlen(start_time) - 1,2))
        flag = 0
        while( flag = 0 )
          if file_exists("pay.txt") != 0
            flag = 1
          else if ss - num(substr(gettime,strlen(gettime) - 4,2)) * 60 + num(substr(gettime,strlen(gettime) - 1,2)) > 5
            flag = 5
          end if
        wend
        if flag = 5
          if file_exists("pay.txt") = 0
            w_message("�訡��!","� �祭�� 5 ᥪ㭤 �� �� ����祭 �⢥� �� ᠩ�! �믮������ �ணࠬ�� ��⠭������!",0)
            end
          else
            fromfile("pay.txt",1)
            while( fromeof = 0 & result = 0 )
              input text
              if strpos("success",text) != 0 then
                result = 1
              end if
            wend
          end if
        else
          if flag = 1 then
            fromfile("pay.txt",1)
            while( fromeof = 0 & result = 0 )
              input text
              if strpos("success",text) != 0 then
                result = 1
              end if
            wend
          end if
        end if
        if result = 0 then
          print "�訡��!","�� 㤠���� ��।��� �㬬� �� ���� " + id_lead
        end if
end function

        klients = keyarray
        if find(exp,1,1,num(nomer)) = 0
          if strlen(trim('��⠐�����'.exp)) = 0
            flag_kl = 0
            count = 0
            w_progress("���� ������ �� �����⠬...",0,size(expkl))
            if findge(expkl,1,1,num(nomer)) = 0 then
              while( eof(expkl) = 0 & '�����'.expkl = num(nomer) )
                count = count + 1
                w_progress(count)
                if find(ok,1,1,'���'.expkl) = 0
                  if strlen(trim('����䮭'.ok)) != 0 | strlen(trim('Email'.ok)) != 0
                    stroka_phone = ""
                    phones = ""
                    for i = 1 to strlen(trim('����䮭'.ok))
                      char = substr(trim('����䮭'.ok),i,1)
                      if char = "," | char = "+" | char = ";" | ( char >= "0" & char <= "9" ) then
                        stroka_phone = stroka_phone + char
                      end if
                    next
                    sp = regex_split("[,;]",stroka_phone)
                    for i = 1 to sp[0]
                      if strlen(trim(sp[i])) != 0 then
                        if strlen(trim(sp[i])) = 7
                          if strlen(trim(phones)) != 0
                            phones = phones + "," + "7351" + trim(sp[i])
                          else
                            phones = phones + "7351" + trim(sp[i])
                          end if
                        else if strlen(trim(sp[i])) = 10 & substr(sp[i],1,1) = "9"
                          if strlen(trim(phones)) != 0
                            phones = phones + ",7" + trim(sp[i])
                          else
                            phones = phones + "7" + trim(sp[i])
                          end if
                        else if regex_match("\+7[0-9]{10}",sp[i]) = 1 & strlen(trim(sp[i])) = 12
                          if strlen(trim(phones)) != 0
                            phones = phones + "," + substr(sp[i],2,11)
                          else
                            phones = phones + substr(sp[i],2,11)
                          end if
                        else if regex_match("8[0-9]{10}",sp[i]) = 1 & strlen(trim(sp[i])) = 11
                          if strlen(trim(phones)) != 0
                            phones = phones + ",7" + substr(trim(sp[i]),2,10)
                          else
                            phones = phones + "7" + substr(trim(sp[i]),2,10)
                          end if
                        end if
                      end if
                    next
                    fl_phones = 0
                    fl_email = 0
                    if strlen(trim(phones)) != 0 then
                      fl_phones = 1
                    end if
                    if strpos("�⪠�",'Email'.ok) = 0 & strlen(trim('Email'.ok)) != 0 then
                      fl_email = 1
                    end if
                    if fl_phones = 1 | fl_email = 1
                      if keyfind(klients,'���'.expkl) = 0
                        klients['���'.expkl] = array(10)
                        klients['���'.expkl][1] = '������'.ok
                        klients['���'.expkl][2] = '�㬬�'.expkl
                        klients['���'.expkl][3] = ""                       rem ID_lead
                        klients['���'.expkl][4] = ""                       rem �����
                        klients['���'.expkl][5] = phones                   rem ⥫�䮭
                        if fl_email = 1
                          klients['���'.expkl][6] = trim('Email'.ok)         rem email
                        else
                          klients['���'.expkl][6] = ""         rem email
                        end if
                        flag_kl = flag_kl + 1
                      else
                        w_message("�訡��!","������ � ����� " + '���'.expkl + " ����砥��� ����� ������ ࠧ�!",0)
                      end if
                    else
                      w_message("�訡��!","� ������ " + trim('������'.expkl) + " ��� ����� �⮢��� ⥫�䮭� � ���஭��� �����! ������ �� �㤥� ����祭 � ��।��� ������!",0)
                    end if
                  else
                    w_message("�訡��!","� ������ " + trim('������'.expkl) + " ��� ����� �⮢��� ⥫�䮭� � ���஭��� �����! ������ �� �㤥� ����祭 � ��।��� ������!",0)
                  end if
                else
                  w_message("�訡��!","� �ࠢ�筨�� �����⮢ �� ������ ������ " + trim('������'.expkl) + " � ����� " + '���'.expkl + "! ������ �� �㤥� ����祭 � ��।��� ������!",0)
                end if
                getnext(expkl)
              wend
            end if
            w_progress

            if flag_kl > 0
              autorization

              count = 0
              w_progress("��।�� �㬬 � CRM...",0,flag_kl)

              s = keyfirst(klients)
              while( strlen(trim(s)) != 0 )
                count = count + 1
                w_progress(count)
                id_contact = ""
                id_lead = ""
                findcontact(klients[s][5],klients[s][6])
                findleads
                if strlen(trim(id_contact)) = 0 | strlen(trim(id_lead)) = 0
                  if find(expkl,2,1,num(nomer),s) = 0 then
                    '���_ID'.expkl = id_lead
                    '�����'.expkl = "�� ������ � CRM"
                    update(expkl)
                  end if
                else
                  yesno = send_sum(klients[s][2])
                  klients[s][3] = id_contact                       rem ID_lead
                  klients[s][4] = id_lead                          rem �����
                  if find(expkl,2,1,num(nomer),s) = 0 then
                    '���_ID'.expkl = id_lead
                    if yesno = 1
                      '�����'.expkl = "��ࠢ����"
                    else
                      '�����'.expkl = "�� ��ࠢ����"
                    end if
                    update(expkl)
                  end if
                end if
                s = keynext(klients,s)
              wend
              w_progress

              '��⠐�����'.exp = tectime
              update(exp)
            else
              w_message("�訡��!","��� �����⮢ ��� ��।��!",0)
            end if
          else
            w_message("�訡��!","��������� 㦥 �믮�����!",0)
          end if
        else
          w_message("�訡��!","�� ������ ���㬥�� � ����஬ " + nomer + "! �믮������ �ணࠬ�� ��ࢠ��!",0)
        end if

        end
