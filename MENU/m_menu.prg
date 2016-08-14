PROCEDURE m_menu
SET SYSMENU TO

DEFINE PAD mmenu_1 OF _MSYSMENU PROMPT '\<ÈÍÔÎÐÌÀÖÈß ÎÒ ËÏÓ' COLOR SCHEME 3 KEY ALT+A, ""
DEFINE PAD mmenu_5 OF _MSYSMENU PROMPT '\<ÑÅÐÂÈÑ' COLOR SCHEME 3 KEY ALT+F , ""
ON PAD mmenu_1 OF _MSYSMENU ACTIVATE POPUP popInfFrLpu
ON PAD mmenu_5 OF _MSYSMENU ACTIVATE POPUP popTuneUp

DEFINE POPUP popInfFrLpu MARGIN RELATIVE shadow COLOR SCHEME 4
DEFINE BAR 01 OF popInfFrLpu PROMPT 'ÏÐÈÅÌ ÈÍÎÐÌÀÖÈÈ ÎÒ ËÏÓ (ÀÈÑ ÎÌÑ)'
DEFINE BAR 02 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 03 OF popInfFrLpu PROMPT 'ÄÅÔÅÊÒÍÛÅ ÈÍÔÎÐÌÀÖÈÎÍÍÛÅ ÏÎÑÛËÊÈ'
DEFINE BAR 04 OF popInfFrLpu PROMPT 'ÏÎÂÒÎÐÍÛÅ ÈÍÔÎÐÌÀÖÈÎÍÍÛÅ ÏÎÑÛËÊÈ'
DEFINE BAR 05 OF popInfFrLpu PROMPT '\-'
DEFINE BAR 06 OF popInfFrLpu PROMPT 'ÂÛÕÎÄ'

ON SELECTION BAR 01 OF popInfFrLpu DO FORM MailView
ON SELECTION BAR 03 OF popInfFrLpu DO FORM MailTrash
ON SELECTION BAR 04 OF popInfFrLpu DO FORM MailDouble
ON SELECTION BAR 06 OF popInfFrLpu CLEAR EVENTS 

DEFINE POPUP popTuneUp MARGIN RELATIVE SHADOW COLOR SCHEME 4
DEFINE BAR 01 OF popTuneUp PROMPT 'ÂÛÁÎÐ ÎÒ×ÅÒÍÎÃÎ ÏÅÐÈÎÄÀ' 
DEFINE BAR 02 OF popTuneUp PROMPT 'ÍÀÑÒÐÎÉÊÀ ÐÀÁÎ×ÈÕ ÄÈÐÅÊÒÎÐÈÉ'
DEFINE BAR 03 OF popTuneUp PROMPT 'ÏÀÐÀÌÅÒÐÛ ÏÅ×ÀÒÈ' 
DEFINE BAR 04 OF popTuneUp PROMPT '\-'
DEFINE BAR 05 OF popTuneUp PROMPT 'ÏÅÐÅÈÍÄÅÊÑÀÖÈß ÁÄ ÍÑÈ'

ON SELECTION BAR 01 OF popTuneUp DO FORM SetPeriod
ON SELECTION BAR 02 OF popTuneUp DO FORM TuneBase
ON SELECTION BAR 03 OF popTuneUp goApp.doForm('set_print','settings')
ON SELECTION BAR 05 OF popTuneUp DO ComReind

SET SYSMENU AUTOMATIC
SET SYSMENU ON