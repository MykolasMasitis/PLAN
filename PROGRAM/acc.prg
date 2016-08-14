FUNCTION Acc(m.mcod)
 m.ppath = pbase + '\' + m.gcperiod + '\' + m.mcod 
 IF !fso.FolderExists(m.ppath)
  RETURN 
 ENDIF 
 IF !fso.FileExists(m.ppath+'\plan.dbf')
  RETURN 
 ENDIF 
 IF OpenFile(m.ppath+'\plan', 'plan', 'shar')>0
  IF USED('plan')
   USE IN plan
  ENDIF 
  SELECT aisoms
  RETURN 
 ENDIF 
 
 SELECT plan 
 IF otm = '0'
  m.acpt = .t.
  REPLACE otm WITH '1'
 ELSE 
  m.acpt = .f.
  REPLACE otm WITH '0'
 ENDIF 
 USE IN plan 
 
 SELECT aisoms
 REPLACE acpt WITH m.acpt
 
 mailview.grid1.refresh


RETURN 