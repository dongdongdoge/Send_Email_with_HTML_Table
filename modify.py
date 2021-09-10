import re
import smtplib
import logging
import datetime
import pandas as pd
import numpy as np
import pymssql
from email.mime.text import MIMEText
from email.header import Header
from pyquery import PyQuery as pq
from dateutil.relativedelta import relativedelta
import configparser
import logging
from logging import handlers

sendFlag = True

class Logger(object):
    level_relations = {
        'debug':logging.DEBUG,
        'info':logging.INFO,
        'warning':logging.WARNING,
        'error':logging.ERROR,
        'crit':logging.CRITICAL
    }

    def __init__(self,filename,level='info',when='midnight',backCount=7,fmt='%(asctime)s [line:%(lineno)d] - %(levelname)s: %(message)s'):
        self.logger = logging.getLogger(filename)
        format_str = logging.Formatter(fmt)#设置日志格式
        self.logger.setLevel(self.level_relations.get(level))#设置日志级别
        sh = logging.StreamHandler()#往屏幕上输出
        sh.setFormatter(format_str) #设置屏幕上显示的格式
        th = handlers.TimedRotatingFileHandler(filename=filename,interval=1,when=when,backupCount=backCount,encoding='utf-8')#往文件里写入#指定间隔时间自动生成文件的处理器
        #实例化TimedRotatingFileHandler
        #interval是时间间隔，backupCount是备份文件的个数，如果超过这个个数，就会自动删除，when是间隔的时间单位，单位有以下几种：
        # S 秒
        # M 分
        # H 小时、
        # D 天、
        # W 每星期（interval==0时代表星期一）
        # midnight 每天凌晨
        th.setFormatter(format_str)#设置文件里写入的格式
        self.logger.addHandler(sh) #把对象加到logger里
        self.logger.addHandler(th)

log = Logger('./log/YTT.log',level='info')

class MyParser(configparser.ConfigParser):
    def as_dict(self):
        d = dict(self._sections)
        for k in d:
            d[k] = dict(d[k])
        return d


def sendMail(body, YYYYWW, mail_list, cc_list):
    smtp_server = r'SMTP.Intel.com'
    from_mail = r'NSG_ZZ_GA@intel.com'
    to_list = mail_list
    cc_list = cc_list
    from_name = r'NSG_ZZ_GA'

    inform = r'{before}.0 - {now}.6'.format(before=str(int(YYYYWW[4:])).zfill(2), now=str(int(YYYYWW[4:])).zfill(2))
    subject = r'SSD Yield & Test Time Report {0}WW{1} ({2}) '.format(YYYYWW[0:4], YYYYWW[4:], inform)

    mail_msg = body + '\n\n' + "Pls contact <a href=\"mailto:wentao.dong@intel.com?cc=jay.ge@intel.com&Subject= Yield and Test Time Report Problems\" target=\"_top\">Administrator</a> for support.   Click <a href=\"http://shwdensgapp1.ccr.corp.intel.com/Mapping/Mapping/Index\"> Here </a> to Map Products. <a href=\"http://shwdensgapp1.ccr.corp.intel.com/Mapping/Mapping/Index2\">Here</a> to Quick Check Model_String. <a href=\"http://shwdensgapp1.ccr.corp.intel.com/YTThistory/YTTtrace/Index\">Here</a> to View History Data Trend."

    message = MIMEText(mail_msg, 'html', 'utf-8')
    message['From'] = Header(from_name)
    message['To'] = Header(';'.join(to_list))
    message['Cc'] = Header(';'.join(cc_list))
    message['Subject'] = Header(subject)

    try:
        log.logger.info(" --- Mail begin!")
        s = smtplib.SMTP()
        s.connect(smtp_server, '25')
        s.sendmail(from_mail, to_list + cc_list, message.as_string())
        s.quit()
        log.logger.info(" --- Mail begin!")
    except smtplib.SMTPException as e:
        log.logger.error(str(e))
        raise e


def tableStyle(htm):
    cfg = MyParser()
    cfg.read('./config.ini', encoding='utf-8')
    family_type = cfg.as_dict()["Family_Type"]
    nand_station = cfg.as_dict()["NAND"]
    optane_station = cfg.as_dict()["OPTANE"]
    for p in nand_station:
        nand_station[p] = eval(nand_station[p])
    for q in optane_station:
        optane_station[q] = eval(optane_station[q])
    table = pq(htm)
    tr = table("tbody tr")
    td = tr.find("td")
    for j in range(-10, 2):
        td.eq(10 + j).addClass("header")
    for i in range(22, len(td), 12):
        if (td.eq(i).text() == '' and td.eq(i - 10).text() == ''):
            for j in range(-10, 2):
                td.eq(i + j).addClass("border")
        else:
            if (td.eq(i + 1).text() == ''):
                continue
            else:
                if (td.eq(i).text()):
                    p = float(td.eq(i + 1).text())
                    td.eq(i + 1).text(str(p))
                    val = float(td.eq(i).text()) - float(td.eq(i + 1).text())
                    bad = (abs(val) > 0.1) and (abs(val) * 10 > float(td.eq(i + 1).text()))
                    if (bad and val < 0):
                        td.eq(i + 1).addClass("high")
                        t = td.eq(i + 1).text()
                        td.eq(i + 1).text(t + ' (+ {:.0%})'.format(abs(val) / float(td.eq(i + 1).text())))
                    if (bad and val > 0):
                        td.eq(i + 1).addClass("low")
                        t = td.eq(i + 1).text()
                        td.eq(i + 1).text(t + ' (- {:.0%})'.format(abs(val) / float(td.eq(i + 1).text())))

    for i in range(7, len(td), 12):
        if (td.eq(i).text() == ''):
            continue
        else:
            family = td.eq(i - 7).text().lower()
            station = td.eq(i - 3).text().lower()
            if (family_type[family] in ("NAND", "DIMM")):
                y = float(td.eq(i).text()[:-1])
                if (y < 95):
                    td.eq(i).addClass("ylow")
                elif (95 <= y < nand_station[station][0]):
                    td.eq(i).addClass("ymid")
                else:
                    td.eq(i).addClass("yhigh")
                y = float(td.eq(i + 1).text()[:-1])
                if (y < 98):
                    td.eq(i + 1).addClass("ylow")
                elif (98 <= y < nand_station[station][1]):
                    td.eq(i + 1).addClass("ymid")
                else:
                    td.eq(i + 1).addClass("yhigh")

            if (family_type[family] == "OPTANE"):
                y = float(td.eq(i).text()[:-1])
                if (y < 95):
                    td.eq(i).addClass("ylow")
                elif (95 <= y < optane_station[station][0]):
                    td.eq(i).addClass("ymid")
                else:
                    td.eq(i).addClass("yhigh")
                y = float(td.eq(i + 1).text()[:-1])
                if (y < 98):
                    td.eq(i + 1).addClass("ylow")
                elif (98 <= y < optane_station[station][1]):
                    td.eq(i + 1).addClass("ymid")
                else:
                    td.eq(i + 1).addClass("yhigh")

            td.eq(i + 3).addClass("TT")
            td.eq(i + 4).addClass("TT")
    return str(table)


def insertNULLRows(df):
    col = df.columns
    k = 0
    p = list(df["PHI_Product"])
    pre = p[0]
    for i in range(len(p)):
        if (p[i] == pre):
            continue
        else:
            pre = p[i]
            df = pd.DataFrame(np.insert(df.values, i + k, values=[None], axis=0))
            k += 1
    df = pd.DataFrame(np.insert(df.values, 0, values=[None], axis=0))
    df.columns = col
    return df.fillna("")


def preFormat(groups, YYYYWW):
    pd.set_option('display.max_colwidth', -1)
    head = """<head>
        <meta charset="utf-8">
        <STYLE TYPE="text/css" MEDIA=screen>
            small{
                color:#8c8c8c;
            }
            table.outline {
                border-collapse: collapse;
                border: 1px solid #000000;
                background-color:#000000;
                width : 90%;
                text-align: center;
                align: center;
            }

            table.outline tr{
                text-align: center;
            }

            table.outline td{
                text-align: center;
            }

            table.dataframe {
                border-collapse: collapse;
                background-color:#ffffff;
                width : 100%;
            }
            table.ins {
                width : 90%;
            }            

            table.dataframe thead {
                border: 2px solid #000000;
                background: #f1f1f1;
                color: #333333;
            }

            table.dataframe tbody {
                border: 2px solid #000000;
            }

            table.dataframe tr {
                font-size: 12px;               
            }

            table.dataframe th {
                vertical-align: bottom;
                font-size: 14px;
                border: 1px solid #000000;
                color: #105de3;
                font-family: arial;
                text-align: center;
                background: #f1f1f1;
            }

            table.dataframe td {
                text-align: center;
                border: 1px solid #000000;
            }     

            body {
                font-family: arial;
                width : 80%;
            }

            h1 {
                left: 100px;
            }

            h4 {
                color: #e10092;
                font-family: arial;
                font-size: 20px;
            }

            td.TT{
                background-color:#d1e0e0;
            }

            td.high{
                background-color:#ff4d4d;
                color: #ffffff;
            }
            td.low{
                background-color:#ffff66;
            }

            td.ylow{
                background-color:#ff4d4d;
                color: #ffffff;
            }
            td.ymid{
                background-color:#ffff66;
            }
            td.yhigh{
                background-color:#d9ffb3;
            }
            td.border{
                background-color:#000000;
            } 

            b.high{
                color:#ff80df;
            }
            b.low{
                color:#ffff66;
            }
            .ins{
                font-size: 12px;
            }
            td
            {
                white-space: nowrap;
            }
            .header{
                background-color:#000000;
            }         
        </STYLE>
    </head>"""
    table = ""
    for name, group in groups:
        group = insertNULLRows(group)
        group = group.fillna('')
        df_html = group.to_html(escape=False, index=False)
        df_html = tableStyle(df_html)
        df_html = '''
            <table class="outline">
        <tr>
            <td>
        ''' + df_html + '''
                </td>
            </tr>
        </table>
        '''
        df_html = "<h4>{0}</h4>".format(name) + df_html
        table += df_html
        table += r"""
        <br>
        """
    inform = r'{before}.0 - {now}.6'.format(before=str(int(YYYYWW[4:])).zfill(2), now=str(int(YYYYWW[4:])).zfill(2))
    body = r"""<body>
        <div class="content">
            <h1> SSD Yield & Test Time Report {YYYY}WW{WW}  <small>({inf})</small> </h1>
        <hr>
            <div>
                <p><b>Instruction : </b> </p>
<table  class="ins">
    <tr>
        <td> <b>Test Time : </b></td>
        <td class="high" style="width:40px"></td>
        <td> : forecast &gt real (delta > 10% and delta > 0.1h) </td>
        <td class="low" style="width:40px"></td>
        <td> : forecast &lt real (delta > 10% and delta > 0.1h) </td>
        <td style="width:40px"></td>
        <td></td>

    </tr>
    <tr>
        <td> <b>NAND/DIMM Yield: </b></td>
        <td class="ylow" style="width:40px"></td>
        <td> : 1A &isin;(0,95%) <b>or</b> Final &isin; (0,98%)</td>
        <td class="ymid" style="width:40px"></td>
        <td> : 1A &isin; [95%,98%) <b>or</b> Final &isin; [98%,99.5%) </td>
        <td class="yhigh" style="width:40px"></td>
        <td> : 1A &isin; [98%,100%] <b>or</b> Final &isin; [99.5%,100%] </td>
        <td>(F1,BI,FT,FT2,FT3) </td>
    </tr>
    <tr>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td> : 1A &isin; [95%,99.5%) <b>or</b> Final &isin; [98%,99.95%) </td>
        <td></td>
        <td> : 1A &isin; [99.5%,100%] <b>or</b> Final &isin; [99.95%,100%] </td>
        <td>(FOQM,OQM,OQME,PFOQM) </td>
    </tr>
    <tr>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td> : 1A &isin; [95%,99%) <b>or</b> Final &isin; [98%,99.8%) </td>
            <td></td>
            <td> : 1A &isin; [99%,100%] <b>or</b> Final &isin; [99.8%,100%] </td>
            <td>(FTCTO) </td>
        </tr>
    <tr>
        <td> <b>OPTANE Yield: </b></td>
        <td class="ylow" style="width:40px"></td>
        <td> : 1A &isin;(0,95%) <b>or</b> Final &isin; (0,98%)</td>
        <td class="ymid" style="width:40px"></td>
        <td> : 1A &isin; [95%,98%) <b>or</b> Final &isin; [98%,99.5%) </td>
        <td class="yhigh" style="width:40px"></td>
        <td> : 1A &isin; [98%,100%] <b>or</b> Final &isin; [99.5%,100%] </td>
        <td>(F1,BI) </td>
    </tr>
    <tr>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td> : 1A &isin; [95%,99%) <b>or</b> Final &isin; [98%,99.5%) </td>
            <td></td>
            <td> : 1A &isin; [99%,100%] <b>or</b> Final &isin; [99.5%,100%] </td>
            <td>(FT,FT2,FT3) </td>
        </tr>
        <tr>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td> : 1A &isin; [95%,99.5%) <b>or</b> Final &isin; [98%,99.95%) </td>
                <td></td>
                <td> : 1A &isin; [99.5%,100%] <b>or</b> Final &isin; [99.95%,100%] </td>
                <td>(FOQM,OQM,OQME,PFOQM) </td>
            </tr>
            <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td> : 1A &isin; [95%,99%) <b>or</b> Final &isin; [98%,99.8%) </td>
                    <td></td>
                    <td> : 1A &isin; [99%,100%] <b>or</b> Final &isin; [99.8%,100%] </td>
                    <td>(FTCTO) </td>
                </tr>
</table>                          

            </div>
        <hr>
            <div>
                <h4></h4>
                {df_html}
            </div>

        <hr>
        </div>

        </body>
        """.format(YYYY=YYYYWW[0:4], WW=YYYYWW[4:], inf=inform, df_html=table)
    html_msg = "<html>" + head + body + "</html>"
    ori_header = r"""<thead>    <tr style="text-align: right;">      <th>Family</th>      <th>PHI_Product</th>      <th>Model_String</th>      <th>Site</th>      <th>Station</th>      <th>Tester</th>      <th>Volume</th>      <th>1A_Yield</th>      <th>FinalYield</th>      <th>Forecast_Yield</th>      <th>Test_Time ( h )</th>      <th>Forecast TT ( h )</th>    </tr>  </thead>"""
    to_header = r"""
    <thead>    
                    <tr>      
                        <th rowspan="2">Family</th>      
                        <th rowspan="2">PHI Product</th>
                        <th rowspan="2">Model String</th>      
                        <th rowspan="2">Site</th>      
                        <th rowspan="2">Station</th>      
                        <th rowspan="2">Tester</th>      
                        <th rowspan="2">Volume</th>      
                        <th colspan="3">Yield</th>      
                        <th colspan="2">Test Time ( h )</th>      
                    </tr>
                    <tr>           
                            <th>1A</th>      
                            <th>Final</th>      
                            <th>Forecast</th>      
                            <th>Real (90%)</th>      
                            <th>Forecast</th>    
                    </tr>
    """
    html_msg = html_msg.replace('\n', '')
    html_msg = html_msg.replace('\t', '')
    html_msg = html_msg.replace(ori_header, to_header)
    html_msg = html_msg.replace('table border="1"', 'table ')
    return html_msg


def sort_Data(df):
    station = set(df['Station'])
    station_sorted = ['PT', 'FT1', 'BI', 'FT', 'FT2', 'FT3', 'FT4', 'FOQM', 'OQM', 'OQME', 'PFOQM', 'FTCTO', 'RMA',
                      'RMAODM']
    station_sorted.extend(list(station - set(station_sorted)))
    df['Station'] = df['Station'].astype('category').cat.set_categories(station_sorted)
    return df.sort_values(['Family', 'PHI_Product', 'Site', 'Station', 'Tester'], ascending=[True] * 5)


def htmlGene(YYYYWW, YYMM):
    try:
        log.logger.info(' -- html generating!')
        conn = pymssql.connect(host='SHWDE9433.CCR.CORP.INTEL.COM', user='FOQM', password='ssd@intel@123',
                            database='MiscDB')
        sql = r'''
        USE [MiscDB]
        EXEC	[dbo].[YieldTTReport]
                @YYMM = N'{0}',
                @YYYYWW = N'{1}'
        '''.format(YYMM, YYYYWW)
        cur = conn.cursor()
        cur.execute(sql)
        df = pd.read_sql(r'''
    SELECT [Family]
        ,[PHI_Product]
        ,[Model_String]
        ,[Site]
        ,[Station]
        ,[Tester]
        ,[Volume]
        ,[1A_Yield]
        ,[FinalYield]
        ,[Forecast_Yield]
        ,[Test_Time ( h )]
        ,[Forecast TT ( h )]
    FROM [MiscDB].[dbo].[Actual_PHITTReport]
    WHERE Station <> 'FT1RESET'
    AND Station <> 'RMA'
    AND Station <> 'RMAODM'
    ORDER BY Family,PHI_Product,Site,Station,Tester
    ''', con=conn)

        # df = df.replace("Kingston","KTC")
        # df = df.replace("KingstonTW","KTCTW")
        df = sort_Data(df)
        groups = df.groupby("Family")
        html = preFormat(groups, YYYYWW)
        conn.close()
        log.logger.info(' -- html generating done!')
        return html
    except Exception as e:
        log.logger.error(str(e))
        raise e


def writeHtmltoFile(html):
    doc = open(r'C:\inetpub\wwwroot\YTThistory\history\{0}WW{1}.html'.format(YYYYWW[0:4],YYYYWW[4:]),'w')
    print(str(html),file=doc)
    doc.close()


def getWW():
    today = datetime.datetime.now()
    lastSunday = today - datetime.timedelta(days=7 + int(today.weekday()))
    lastSunday = lastSunday.strftime('%Y%m%d')
    sql = '''
    SELECT [WW]
    FROM [MiscDB].[dbo].[Dim_Date]
    where Date = '{0}'
    '''.format(lastSunday)
    cnxx = pymssql.connect(host='SHWDE9433.CCR.CORP.INTEL.COM', user='FOQM', password='ssd@intel@123',
                           database='MiscDB')
    cur = cnxx.cursor()
    cur.execute(sql)
    result = cur.fetchall()
    return str(result[0][0][0:4]) + str(result[0][0][-2:])


def getMM():
    today = datetime.datetime.now()
    dt = today - datetime.timedelta(days=7)
    return datetime.datetime.strftime(dt, '%Y-%m')[2:]


if __name__ == "__main__":
    try:
        YYYYWW = getWW()
        YYMM = getMM()
        # cfg1 = MyParser()
        # cfg1.read('./config.ini',encoding='utf-8')
        # mails = cfg1.as_dict()["MAIL_LIST"]
        # mail_list = eval(mails['r'])
        # cc_list = eval(mails['cc'])

        mail_list = [r'wentao.dong@intel.com']
        cc_list = [r'836347620@qq.com']
        html = htmlGene(YYYYWW,YYMM)
        #html = 'hello this is world'
        writeHtmltoFile(html)
        sendMail(html,YYYYWW,mail_list,cc_list)
    except Exception as e:
        log.logger.error(str(traceback.format_exc()))


