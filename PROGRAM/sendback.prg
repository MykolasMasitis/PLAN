FUNCTION SendBack(para1,para2,para3,para4)

 m.lpuid    = para1
 m.mcod     = para2
 m.cfrom    = para3
 m.cmessage = para4
 
 m.cTO  = IIF(!EMPTY(ALLTRIM(cfrom)), ALLTRIM(cfrom), ;
  IIF(SEEK(lcLpuID, 'sprabo', 'lpu_id'), 'OMS@'+ALLTRIM(sprabo.abn_name), ''))
 
 m.mmy = SUBSTR(gcPeriod,5,2)+SUBSTR(gcPeriod,4,1)

 m.un_id    = SYS(3)
 m.bansfile = 'b' + m.un_id
 m.tansfile = 't' + m.un_id
 m.dfile    = 'd' + m.un_id
 m.mmid     = m.un_id+'@'+m.qmail
 m.csubj    = 'PLAN#'+STR(m.lpuid,4)+'#'+UPPER(m.qcod)+'#'+DTOS(DATE())

 poi = fso.CreateTextFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.tansfile)

 poi.WriteLine('To: '+m.cTO)
 poi.WriteLine('Message-Id: ' + m.mmid)
 poi.WriteLine('Subject: ' + m.csubj)
 poi.WriteLine('Content-Type: multipart/mixed')
 poi.WriteLine('Resent-Message-Id: '+ALLTRIM(m.cmessage))
 poi.WriteLine('Attachment: '+m.dfile+' PLANP_'+STR(m.lpuid,4)+'.'+m.mmy)
 
 poi.Close
 
 fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\plan.dbf', pAisOms+'\oms\output\'+m.dfile)
 fso.CopyFile(pbase+'\'+m.gcperiod+'\'+m.mcod+'\'+m.tansfile, pAisOms+'\oms\output\'+m.bansfile)

RETURN 