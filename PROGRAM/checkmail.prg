PROCEDURE CheckMail
PARAMETERS lcUser

IF !IsAisDir() && Проверка наличия директорий, OMS, INPUT, OUTPUT
 RETURN 
ENDIF 

oMailDir        = fso.GetFolder(pAisOms+'\&lcUser\input')
MailDirName     = oMailDir.Path
oFilesInMailDir = oMailDir.Files
nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('ОБНАРУЖЕНО '+ALLTRIM(STR(nFilesInMailDir))+' ФАЙЛОВ!', 0+64, lcUser)

IF nFilesInMailDir<=0
 RETURN 
ENDIF 

IF OpenTemplates() != 0
 =CloseTemplates() 
 RETURN 
ENDIF 

WAIT "ПРОСМОТР ПОЧТЫ..." WINDOW NOWAIT 
SELECT AisOms
prvorder = ORDER('aisoms')
SET ORDER TO 

m.un_id = SYS(3)

OldEscStatus = SET("Escape")
SET ESCAPE OFF 
CLEAR TYPEAHEAD 

FOR EACH oFileInMailDir IN oFilesInMailDir

 SCATTER MEMVAR BLANK
 m.mmy   = SUBSTR(m.gcPeriod,5,2)+SUBSTR(m.gcPeriod,4,1)

 m.BFullName = oFileInMailDir.Path
 m.bname     = oFileInMailDir.Name
 m.recieved  = oFileInMailDir.DateLastModified
 m.lpuid     = 0
 m.processed = DATETIME()
 
 m.cfrom      = ''
 m.cdate      = ''
 m.cmessage   = ''
 m.resmesid   = ''
 m.csubject   = ''
 m.csubject1  = ''
 m.csubject2  = ''
 m.attachment = ''
 m.bodypart   = ''

 m.attaches   = 0 && Сколько присоединенных файлов в одной ИП
 DIMENSION dattaches(10,2)
 dattaches = ''

 m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
 DIMENSION dbparts(10,2)
 dbparts = ''

 DO CASE 
 CASE LOWER(oFileInMailDir.Name) = 'b'

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 =FCLOSE(CFG)

 IF UPPER(LEFT((ALLTRIM(m.csubject)),4)) != 'PLAN'
  LOOP 
 ENDIF 
 IF RIGHT(ALLTRIM(dattaches(1, 2)),3) != m.mmy
  LOOP 
 ENDIF 

 m.sent = dt2date(m.cdate)
* MESSAGEBOX('"'+m.cdate+'"',0+64,'')
 WAIT m.cfrom WINDOW NOWAIT 
   
 m.llIsSubject = .F.

 m.adresat = PADR(LOWER(SUBSTR(m.cfrom,AT('@',m.cfrom)+1)),27)
 IF m.cfrom='oms@spuemias.msk.oms'
  m.lpuid = INT(VAL(SUBSTR(m.csubject,AT('#',m.csubject,2)+1,AT('#',m.csubject,3)-AT('#',m.csubject,2)-1)))
 ELSE 
  m.lpuid   = IIF(SEEK(m.adresat, "sprabo", "abn_name"), sprabo.object_id, m.lpuid)
 ENDIF 
 
 IF m.lpuid==0
*  MESSAGEBOX('АДРЕСАТ '+UPPER(ALLTRIM(m.adresat))+' НЕ НАЙДЕН В СПРАВОЧНИКЕ SPRABOXX.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 
 
 m.mcod    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.mcod, "")

 IF EMPTY(m.mcod)
*  MESSAGEBOX('АДРЕСАТ '+UPPER(ALLTRIM(m.adresat))+' НЕ НАЙДЕН В СПРАВОЧНИКЕ SPRLPUXX.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 

 m.cokr    = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.cokr, "")
 m.moname  = IIF(SEEK(m.lpuid, "sprlpu"), sprlpu.name, "")
 m.usr     = IIF(SEEK(m.lpuid, "usrlpu"), 'USR'+PADL(usrlpu.usr,3,'0'), "")

 IF EMPTY(m.usr) AND m.gcUser!='OMS'
*  MESSAGEBOX('ЛПУ '+m.mcod+' НЕ "ПРИВЯЗАНО" К ПОЛЬЗВАТЕЛЮ В USRLPU.DBF!',0+48,lcUser)
  LOOP 
 ENDIF 
 
 IF m.usr != m.gcUser AND m.gcUser!='OMS'
  LOOP 
 ENDIF 

 m.IsVed   = IIF(LEFT(m.mcod,1) == '0', .F., .T.)

 m.previous_id = m.un_id
 m.un_id     = SYS(3)
 DO WHILE m.un_id = m.previous_id
  m.un_id     = SYS(3)
 ENDDO 
 m.previous_id = m.un_id

 m.tansfile  = 'tok_' + m.mcod
 m.bansfile  = 'bok_' + m.mcod
 iii = 1
 DO WHILE fso.FileExists(pAisOms+'\OMS\OUTPUT\'+m.bansfile)
  m.tansfile  = 'tok_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bansfile  = 'bok_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 

 m.messageid = ALLTRIM(m.un_id+'.'+m.gcUser+'@'+m.qmail)

 && Присоединено ли что-нибудь к файлу? Если нет, то - в спам!
 IF m.attaches == 0 AND m.bparts == 0 && Если к файлу ничего не присоединено!
  TextToWrite="MyComment: к файлу ничего не присоединено"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  LOOP 
 ENDIF 
 && Присоединено ли что-нибудь к файлу? Если нет, то - в спам!

 && Проверка комплектности посылки
 IsComplect = .T.
 IF m.attaches>0
  FOR natt = 1 TO m.attaches
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    IsComplect = .F.
*    MESSAGEBOX('ПРИСОЕДИНЕННЫЙ К ФАЙЛУ '+m.bname+CHR(13)+CHR(10)+;
     ' ATTACHMENT '+dattaches(natt,1)+ ' ОТСУТСТВУЕТ!', 0+48, lcUser)
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 
 IF IsComplect = .F.
  FOR natt = 1 TO m.attaches
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1), SpamDir+'\'+dattaches(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dattaches(natt,1))
   ENDIF 
  ENDFOR 
  TextToWrite="MyComment: отсутствует или недоступен присоединенный файл"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  LOOP 
 ENDIF  

 IsComplect = .T.
 IF m.bparts>0
  FOR natt = 1 TO m.bparts
   IF !fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    IsComplect = .F.
*    MESSAGEBOX('ПРИСОЕДИНЕННЫЙ К ФАЙЛУ '+m.bname+CHR(13)+CHR(10)+;
     ' DODY-PART '+dbparts(natt,1)+ ' ОТСУТСТВУЕТ!', 0+48, lcUser)
    LOOP 
   ENDIF 
  ENDFOR 
 ENDIF 
 IF IsComplect = .F.
  FOR natt = 1 TO m.bparts
   IF fso.FileExists(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
    fso.CopyFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1), SpamDir+'\'+dbparts(natt,1), .t.)
    fso.DeleteFile(pAisOms+'\'+lcUser+'\input\'+dbparts(natt,1))
   ENDIF 
  ENDFOR 
  TextToWrite="MyComment: отсутствует или недоступен присоединенный файл"
  fso.CopyFile(m.BFullName, SpamDir+'\'+m.bname, .t.)
  fso.DeleteFile(m.BFullName)
  =WriteInBFile(SpamDir+'\'+m.bname, TextToWrite)
  LOOP 
 ENDIF  
 && Проверка комплектности посылки

 poi = fso.CreateTextFile(pAisOms+'\&lcUser\output\'+m.tansfile)
 poi.WriteLine('To: '+m.cfrom)
 poi.WriteLine('Message-Id: ' + m.messageid)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: ' + m.cmessage)

 && Если посылка повторная
* IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod'), .t., .f.)
 IsThisIPDouble = IIF(SEEK(m.mcod, 'AisOms', 'mcod') AND !EMPTY(AisOms.Sent), .t., .f.)
 IF IsThisIPDouble
  frst_time = AisOms.Sent
  IF frst_time > m.sent && Принятая ранее посылка отправлена позже обнаруженной!
*   MESSAGEBOX('AisOms.Sent='+DTOC(AisOms.Sent)+'m.sent'+m.cdate,0+64,'frst_time > m.sent')
   m.csubject = m.csubject1 + '99' +m.csubject2
   m.cerrmessage = [Уже загружена более поздяя посылка]
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.bansfile)
   fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pArc+'\IPS\OUTPUT\'+m.bansfile)
   fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
   DoubleDir = pDouble + '\' + m.mcod
   IF !fso.FolderExists(DoubleDir)
    fso.CreateFolder(DoubleDir)
   ENDIF 

   fso.CopyFile(m.BFullName, DoubleDir+'\'+m.bname)
   fso.DeleteFile(m.BFullName)
   IF m.attaches>0
    FOR nattach = 1 TO m.attaches
     m.dname   = ALLTRIM(dattaches(nattach, 1))
     IF !EMPTY(m.dname)
      fso.CopyFile(MailDirName + '\' + m.dname, pDouble+'\'+m.mcod+'\'+m.dname)
      fso.DeleteFile(MailDirName + '\' + m.dname)
     ENDIF 
    ENDFOR 
   ENDIF 
   IF m.bparts > 0
    FOR npart = 1 TO m.bparts
     m.bpname   = ALLTRIM(dbparts(npart, 1))
     IF !EMPTY(m.bpname)
      fso.CopyFile(MailDirName + '\' + m.bpname, pDouble+'\'+m.mcod+'\'+m.dname)
      fso.DeleteFile(MailDirName + '\' + m.bpname)
     ENDIF 
    ENDFOR 
   ENDIF 

   IF NOT SEEK(m.cmessage, "daisoms")
    INSERT INTO daisoms FROM MEMVAR 
   ENDIF

   LOOP 

  ELSE                  && Обнаруженная посылка более свежая, чем принятая ранее!

*   MESSAGEBOX('AisOms.Sent='+DTOC(AisOms.Sent)+'m.sent'+DTOC(m.sent),0+64,'frst_time > m.sent')

   m.prv_bfile = ALLTRIM(AisOms.bname)
   DoubleDir   = pDouble + '\' + m.mcod
   IF !fso.FolderExists(DoubleDir)
    fso.CreateFolder(DoubleDir)
   ENDIF 
   
   CFG = FOPEN(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.prv_bfile)
   DO WHILE NOT FEOF(CFG)
    READCFG = FGETS (CFG)
    IF UPPER(READCFG) = 'ATTACHMENT'
     m.dbl_attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
     m.dbl_dname      = ALLTRIM(SUBSTR(m.dbl_attachment, 1, AT(" ",m.dbl_attachment)-1)) && Название d-файла
     m.dbl_attname    = ALLTRIM(SUBSTR(m.dbl_attachment, AT(" ",m.dbl_attachment)+1))    && Фактическое название файла
     IF !EMPTY(m.dbl_attname)
      fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\plan.dbf', DoubleDir+'\'+m.dbl_dname)
     ENDIF 
    ENDIF 
   ENDDO
   = FCLOSE (CFG)
   fso.CopyFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\'+m.prv_bfile, DoubleDir+'\'+m.prv_bfile, .t.)
   
   fso.DeleteFile(m.pBase+'\'+m.gcPeriod+'\'+m.mcod+'\*.*')

   m.t_BName      = AisOms.BName
   m.t_Sent       = AisOms.Sent
   m.t_Recieved   = AisOms.Recieved
   m.t_Processed  = AisOms.Processed
   m.t_CMessage   = AisOms.CMessage
   m.t_dname      = AisOms.dname
*   DELETE IN AisOms && !!!
   
*   MailView.get_recs.value = MailView.get_recs.value - 1

   MailView.recs   = MailView.recs   - 1
   MailView.s1   = MailView.s1   - aisoms.s1
   MailView.s2   = MailView.s2   - aisoms.s2
   MailView.s3   = MailView.s3   - aisoms.s3
   MailView.s4   = MailView.s4   - aisoms.s4
   MailView.sy   = MailView.sy   - aisoms.sy

   IF NOT SEEK(m.t_CMessage, "daisoms")

    INSERT INTO daisoms (LpuId,Mcod,BName,Sent,Recieved,Processed,CFrom,CMessage,dname) ;
     VALUES ;
     (m.lpuid,m.mcod,m.t_BName,m.t_Sent,m.t_Recieved,m.t_Processed,m.cfrom,m.t_CMessage,m.t_dname)
   ENDIF

   RELEASE m.t_BName,m.t_Sent,m.t_Recieved,m.t_Processed,m.t_CMessage,m.t_dname

  ENDIF 
 ENDIF 
 && Если посылка повторная

 && Если это нормальная и новая посылка!
 InDirPeriod = pBase + '\' + m.gcPeriod
 IF !fso.FolderExists(InDirPeriod)
  fso.CreateFolder(InDirPeriod)
 ENDIF 
 InDir = pBase + '\' + m.gcPeriod + '\' + m.mcod
 IF !fso.FolderExists(InDir)
  fso.CreateFolder(InDir)
 ENDIF 

 fso.CopyFile(m.BFullName, InDir + '\' + m.bname)
 fso.CopyFile(m.BFullName, pArc + '\IPS\INPUT\' + m.bname)
 fso.DeleteFile(m.BFullName)
 IF m.attaches>0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
*   m.aattname = SUBSTR(m.aattname,1,IIF(AT('.',m.aattname)>0,AT('.',m.aattname)-1,LEN(ALLTRIM(m.aattname))))+'.zip'
*   m.aattname = SUBSTR(m.aattname,1,IIF(AT('.',m.aattname)>0,AT('.',m.aattname)-1,LEN(ALLTRIM(m.aattname))))
*   MESSAGEBOX(m.dname+CHR(13)+CHR(10)+m.ddname+CHR(13)+CHR(10)+m.aattname,0+64,'')
   IF !EMPTY(m.ddname)
    fso.CopyFile(MailDirName + '\' + m.ddname, InDir+'\plan.dbf', .t.)
    fso.CopyFile(MailDirName + '\' + m.ddname, parc+'\IPS\INPUT\'+m.ddname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
   ENDIF 
  ENDFOR 
 ENDIF 
 IF m.bparts > 0
  FOR npart = 1 TO m.bparts
   m.bpname   = ALLTRIM(dbparts(npart, 1))
   IF !EMPTY(m.bpname)
    fso.CopyFile(MailDirName + '\' + m.bpname, InDir+'\'+m.bpname, .t.)
    fso.CopyFile(MailDirName + '\' + m.bpname, parc+'\IPS\INPUT\'+m.bpname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.bpname)
   ENDIF 
  ENDFOR 
 ENDIF 

 =OpenFile(Indir+'\plan',  "plan",  "SHARED")
 
* MESSAGEBOX('OK',0+64,'')
 
 IF !CheckFilesStucture()
  LOOP 
 ENDIF 
 
 m.csubject = m.csubject1 + '01' +m.csubject2
 poi.WriteLine('Subject: '+m.csubject)
 poi.WriteLine('BodyPart: OK' )
 poi.Close

 fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pAisOms+'\'+lcUser+'\OUTPUT\'+bansfile)
 fso.CopyFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile, pArc+'\IPS\OUTPUT\'+bansfile)
 fso.DeleteFile(pAisOms+'\'+lcUser+'\OUTPUT\'+tansfile)
 
 SELECT plan
 m.ny = 0
 m.n1 = 0
 m.n2 = 0
 m.n3 = 0
 m.n4 = 0
 m.sy = 0
 m.s1 = 0
 m.s2 = 0
 m.s3 = 0
 m.s4 = 0
 SCAN 
  SCATTER MEMVAR
  m.n1 = m.n1+m.pos_n1+m.pos_p1+m.obr1+m.dn_st1+m.st1+m.eco1+m.vmp1
  m.n2 = m.n2+m.pos_n2+m.pos_p2+m.obr2+m.dn_st2+m.st2+m.eco2+m.vmp2
  m.n3 = m.n3+m.pos_n3+m.pos_p3+m.obr3+m.dn_st3+m.st3+m.eco3+m.vmp3
  m.n4 = m.n4+m.pos_n4+m.pos_p4+m.obr4+m.dn_st4+m.st4+m.eco4+m.vmp4
  m.ny = m.ny + m.n1 + m.n2 + m.n3 + m.n4

  m.s1 = m.s1+m.s_n1+m.s_p1+m.s_o1+m.s_dn_st1+m.s_st1+m.s_eco1+m.s_vmp1
  m.s2 = m.s2+m.s_n2+m.s_p2+m.s_o2+m.s_dn_st2+m.s_st2+m.s_eco2+m.s_vmp2
  m.s3 = m.s3+m.s_n3+m.s_p3+m.s_o3+m.s_dn_st3+m.s_st3+m.s_eco3+m.s_vmp3
  m.s4 = m.s4+m.s_n4+m.s_p4+m.s_o4+m.s_dn_st4+m.s_st4+m.s_eco4+m.s_vmp4
  m.sy = m.sy + m.s1 + m.s2 + m.s3 + m.s4

 ENDSCAN 
 USE IN plan

 MailView.recs   = MailView.recs   + 1
 MailView.s1   = MailView.s1   + m.s1
 MailView.s2   = MailView.s2   + m.s2
 MailView.s3   = MailView.s3   + m.s3
 MailView.s4   = MailView.s4   + m.s4
 MailView.sy   = MailView.sy   + m.sy

 MailView.n1   = MailView.n1   + m.n1
 MailView.n2   = MailView.n2   + m.n2
 MailView.n3   = MailView.n3   + m.n3
 MailView.n4   = MailView.n4   + m.n4
 MailView.ny   = MailView.ny   + m.ny

 =SEEK(m.mcod, 'aisoms', 'mcod')
 MailView.refresh

 UPDATE aisoms SET bname=m.bname, dname=m.dname, sent=m.sent, recieved=m.recieved, ;
  processed=m.processed, cfrom=m.cfrom, cmessage=m.cmessage, ;
  sy=m.sy, s1=m.s1, s2=m.s2, s3=m.s3, s4=m.s4,;
  ny=m.ny, n1=m.n1, n2=m.n2, n3=m.n3, n4=m.n4, acpt=.f.;
  WHERE mcod=m.mcod

 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 CASE LOWER(oFileInMailDir.Name) = 'r' && Если это r-файл

 CFG = FOPEN(m.BFullName)
 =ReadCFGFile()
 =FCLOSE (CFG)

 WAIT m.cfrom WINDOW NOWAIT 
   
 fso.CopyFile(m.BFullName, DaemonDir + '\' + m.bname, .t.)
 fso.DeleteFile(m.BFullName)

 IF m.attaches > 0
  FOR nattach = 1 TO m.attaches
   m.ddname   = ALLTRIM(dattaches(nattach, 1))
   m.aattname = ALLTRIM(dattaches(nattach, 2))
*   m.aattname = SUBSTR(m.aattname,1,IIF(AT('.',m.aattname)>0,AT('.',m.aattname)-1,LEN(ALLTRIM(m.aattname))))+'.zip'
*   m.aattname = SUBSTR(m.aattname,1,IIF(AT('.',m.aattname)>0,AT('.',m.aattname)-1,LEN(ALLTRIM(m.aattname))))
   IF !EMPTY(m.dname) AND fso.FileExists(MailDirName + '\' + m.ddname)
    fso.CopyFile(MailDirName + '\' + m.ddname, DaemonDir+'\'+m.aattname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.ddname)
   ENDIF 
  ENDFOR 
 ENDIF 

 IF m.bparts > 0
  FOR npart = 1 TO m.bparts
   m.bpname   = ALLTRIM(dbparts(npart, 1))
   IF !EMPTY(m.bpname) AND fso.FileExists(MailDirName + '\' + m.bpname)
    fso.CopyFile(MailDirName + '\' + m.bpname, DaemonDir+'\'+m.bpname, .t.)
    fso.DeleteFile(MailDirName + '\' + m.bpname)
   ENDIF 
  ENDFOR 
 ENDIF 

 SELECT AisOms
 
 WAIT CLEAR 
 MailView.Refresh 

 ENDCASE 

 IF CHRSAW(0) 
  IF INKEY() == 27
   IF MESSAGEBOX('ВЫ ХОТИТЕ ПРЕРВАТЬ ОБРАБОТКУ?',4+32,'') == 6
    EXIT 
   ENDIF 
  ENDIF 
 ENDIF 

NEXT && Цикл по файлам

SET ESCAPE &OldEscStatus

=CloseTemplates() 

SET ORDER TO (prvorder)
MailView.Refresh
MailView.LockScreen=.f.

WAIT CLEAR 

nFilesInMailDir = oFilesInMailDir.Count

MESSAGEBOX('ОСТАЛОСЬ '+ALLTRIM(STR(nFilesInMailDir))+' НЕОБРАБОТАННЫХ ИП!', 0+64, lcUser)

RETURN 

FUNCTION CopyToTrash(lcPath, nTip)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pAisOms+'\&lcUser\OUTPUT\'+m.bansfile)
 fso.CopyFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile, pArc+'\IPS\OUTPUT\'+m.bansfile)
 fso.DeleteFile(pAisOms+'\&lcUser\OUTPUT\'+m.tansfile)
 TrashDir = pTrash + '\' + m.mcod
 IF !fso.FolderExists(TrashDir)
  fso.CreateFolder(TrashDir)
 ENDIF 
 fso.CopyFile(lcPath + '\' + m.bname, TrashDir+'\'+m.bname)
 fso.DeleteFile(lcPath + '\' + m.bname)

 FOR nattach = 1 TO m.attaches
  IF !EMPTY(ALLTRIM(dattaches(nattach, 1)))
   fso.CopyFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)), TrashDir + '\' + ALLTRIM(dattaches(nattach, 2)))
   fso.DeleteFile(lcPath + '\' + ALLTRIM(dattaches(nattach, nTip)))
  ENDIF 
 ENDFOR 

 IF NOT SEEK(m.cmessage, "taisoms", "cmessage")
  INSERT INTO taisoms FROM MEMVAR 
 ENDIF
RETURN 

FUNCTION ClDir
 IF fso.FileExists(plan)
  DELETE FILE &plan
 ENDIF 
RETURN 

FUNCTION OpenTemplates
 tn_result = 0
 tn_result = tn_result + OpenFile("&ptempl\planp.mmy", "pl_et", "SHARED")
RETURN tn_result

FUNCTION CloseTemplates
 IF USED('pl_et')
  USE IN pl_et
 ENDIF 
RETURN 

FUNCTION CloseItems
 IF USED('plan')
  USE IN plan
 ENDIF 
RETURN 

FUNCTION CompFields(NameOfFile)
 FOR nFld = 1 TO fld_1
  IF (tabl_1(nFld,1) == tabl_2(nFld,1)) AND ;
     (tabl_1(nFld,2) == tabl_2(nFld,2)) AND ;
     (tabl_1(nFld,3) == tabl_2(nFld,3))
  ELSE 
   =CloseItems()
*   =ClDir()
   m.csubject = m.csubject1 + '08' +m.csubject2
   m.cerrmessage = [Wrong structure of ] + NameOfFile
   IF m.llIsSubject = .F.
    m.llIsSubject = .T.
    poi.WriteLine('Subject: '+m.csubject)
   ENDIF 
   poi.WriteLine('BodyPart: ' + m.cerrmessage)
   poi.Close
   RETURN 0 
  ENDIF 
 ENDFOR 
RETURN 1

FUNCTION DiffFields(NameOfFile)
 =CloseItems()
* =ClDir()
 m.csubject = m.csubject1 + '08' +m.csubject2
 m.cerrmessage = [Wrong number of fields in ] + NameOfFile
 IF m.llIsSubject = .F.
  m.llIsSubject = .T.
  poi.WriteLine('Subject: '+m.csubject)
 ENDIF 
 poi.WriteLine('BodyPart: ' + m.cerrmessage)
 poi.Close
RETURN 

FUNCTION IsAisDir()
 IF !fso.FolderExists(pAisOms)
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms, 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\INPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\INPUT', 0+16, '')
  RETURN .F.
 ENDIF 

 IF !fso.FolderExists(pAisOms+'\&lcUser\OUTPUT')
  MESSAGEBOX('ОТСУТСТВУЕТ ДИРЕКТОРИЯ '+pAisOms+'\&lcUser\OUTPUT', 0+16, '')
  RETURN .F. 
 ENDIF

RETURN .T. 

FUNCTION WriteInBFile(BFullName, TextToWrite)
 CFG = FOPEN(BFullName,12)
 IsMyCommentExists = .F.
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  IF UPPER(READCFG) = 'MYCOMMENT'
   IsMyCommentExists = .T.
   LOOP 
  ENDIF 
 ENDDO
 IF !IsMyCommentExists
  nFileSize = FSEEK(CFG,0,2)
  =FWRITE(CFG, TextToWrite)
 ENDIF 
 = FCLOSE (CFG)
RETURN 

FUNCTION CreateFilesStructure
 CREATE TABLE (PlanP) ;
  (recid c(7), profil c(3), ;
   pos_n1 n(8), s_n1 n(15,2), pos_p1 n(8), s_p1 n(15,2), obr1 n(8), s_o1 n(15,2), dn_st1 n(8), s_dn_st1 n(15,2), st1 n(8), s_st1 n(15,2), eco1 n(8), s_eco1 n(15,2), vmp1 n(8), s_vmp1 n(15,2),;
   pos_n2 n(8), s_n2 n(15,2), pos_p2 n(8), s_p2 n(15,2), obr2 n(8), s_o2 n(15,2), dn_st2 n(8), s_dn_st2 n(15,2), st2 n(8), s_st2 n(15,2), eco2 n(8), s_eco2 n(15,2), vmp2 n(8), s_vmp2 n(15,2),;
   pos_n3 n(8), s_n3 n(15,2), pos_p3 n(8), s_p3 n(15,2), obr3 n(8), s_o3 n(15,2), dn_st3 n(8), s_dn_st3 n(15,2), st3 n(8), s_st3 n(15,2), eco3 n(8), s_eco3 n(15,2), vmp3 n(8), s_vmp3 n(15,2),;
   pos_n4 n(8), s_n4 n(15,2), pos_p4 n(8), s_p4 n(15,2), obr4 n(8), s_o4 n(15,2), dn_st4 n(8), s_dn_st4 n(15,2), st4 n(8), s_st4 n(15,2), eco4 n(8), s_eco4 n(15,2), vmp4 n(8), s_vmp4 n(15,2),;
   datap d, nom_p c(25), otm c(1))
 USE 
RETURN 

FUNCTION OpenLocalFiles
 USE (planp) IN 0 ALIAS planp SHARED
RETURN 

FUNCTION CheckFilesStucture

 fld_1 = AFIELDS(tabl_1, 'plan') && проверка r-файла
 fld_2 = AFIELDS(tabl_2, 'pl_et')
 IF fld_1 == fld_2 && Кол-во полей совпадает!
  FieldsIdent = CompFields('plan') && 0 - есть отличия, 1 - полное совпадение
  IF FieldsIdent==0
   =CopyToTrash(m.InDir,2)
   =ClDir()
   RETURN .F.
  ENDIF 
 ELSE 
  =DiffFields(plan)
  =CopyToTrash(m.InDir,2)
  =ClDir()
  RETURN .F.
 ENDIF 

RETURN .T. 

FUNCTION ReadCFGFile
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  DO CASE
   CASE UPPER(READCFG) = 'FROM'
    m.cfrom = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'DATE'
    m.cdate = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'MESSAGE'
    m.cmessage = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'RESENT-MESSAGE-ID'
    m.resmesid = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'SUBJECT'
    m.csubject = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    m.csubject1 = LEFT(m.csubject, RAT('#',m.csubject,2))   && Делим subject для последующей вставки кода результата
    m.csubject2 = SUBSTR(m.csubject, RAT('#',m.csubject,1)) && Делим subject для последующей вставки кода результата
   CASE UPPER(READCFG) = 'ATTACHMENT'
    m.attaches   = m.attaches + 1
    m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && Название d-файла
    dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && Фактическое название файла
   CASE UPPER(READCFG) = 'BODYPART'
    m.bparts   = m.bparts + 1
    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 