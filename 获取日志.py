#服务器日志打印
import codecs
import os
import sys
import traceback
import win32con
import win32evtlog
import win32evtlogutil
import time
import winerror
import openpyxl
from openpyxl import Workbook

def Get_SecureEvents(logtypes,basepath):
    server="localhost"
    for logtype in logtypes:
        path=os.path.join(basepath,server+'_'+logtype+'_log.log')
        getEventLogs(server,logtype,path)
  

def getEventLogs(server,logtype,logpath):
    print('正在获取'+logtype+'日志')
    log=codecs.open(logpath,encoding='utf-8',mode='w')
#读取本机的系统日志
    hand=win32evtlog.OpenEventLog(server,logtype)
#读取系统日志总行数
    total=win32evtlog.GetNumberOfEventLogRecords(hand)
#按序读取
    flags=win32evtlog.EVENTLOG_BACKWARDS_READ|win32evtlog.EVENTLOG_SEQUENTIAL_READ
##错误级别字典
    evt_dict = {win32con.EVENTLOG_AUDIT_FAILURE: 'EVENTLOG_AUDIT_FAILURE',
                win32con.EVENTLOG_AUDIT_SUCCESS: 'EVENTLOG_AUDIT_SUCCESS',
                win32con.EVENTLOG_INFORMATION_TYPE: 'EVENTLOG_INFORMATION_TYPE',
                win32con.EVENTLOG_WARNING_TYPE: 'EVENTLOG_WARNING_TYPE',
                win32con.EVENTLOG_ERROR_TYPE: 'EVENTLOG_ERROR_TYPE'}

    try:
        events=1
        i=1
        while events:
            events=win32evtlog.ReadEventLog(hand,flags,0)
            for ev_obj in events:
                the_time=ev_obj.TimeGenerated.Format() #'12/23/999 15:54:09'
                evt_id=winerror.HRESULT_CODE(ev_obj.EventID)
                computer=ev_obj.ComputerName
                cat=ev_obj.EventCategory
                record=ev_obj.RecordNumber
                msg=win32evtlogutil.SafeFormatMessage(ev_obj,logtype)
                source=str(ev_obj.SourceName)
                if not ev_obj.EventType in evt_dict.keys():
                    evt_type="unknow"
                else:
                    evt_type=evt_dict[ev_obj.EventType]
                    log.write("Event Date/Time: %s\n" %the_time)
                    log.write("EventID /Type: %s / %s\n" %(evt_id,evt_type))
                    log.write("Record #%s\n" %record)
                    log.write("Source: %s\n\n" %source)
                    log.write(msg)
                    log.write('\n\n')
                    log.write('__________________________________________________')
                    log.write('\n\n')
                    ws.cell(row=i,column=1).value=msg
                    i=i+1
                    
    except:
            print(traceback.print_exc(sys.exc_info()))


    
    
if __name__=="__main__":
    i=1
    logtypes=["Security"]
    wb=Workbook()
    ws=wb.active
    ws.title='日志管理'
    Get_SecureEvents(logtypes,'D:\\')
    wb.save('D:\\1.xlsx')
        
