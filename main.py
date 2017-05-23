import xlsxwriter
import psycopg2
import smtplib
import os
import time
import datetime

class KPI:

    def __init__(self):
        #Carga configuracion.
        self.dbConfiguration = {
            'host': '127.0.0.1',
            'username': 'postgres',
            'port': '65432',
            'db': 'servicedesk'
        }
        self.mailConfiguration = {
            'host': 'smtp.gmail.com',
            'smtpPort':'587',
            'username': 'soporte@openbco.com',
            'to': 'diegomecatronico@gmail.com',
            'cc': 'ingenieria1@openbco.com'
        }
        self.settings = {
            'from': '2017-04-01',
            'to': '2017-04-20'
        }
        self.logLocation = "log.txt"

    def obtainDataFromServiceDesk(self):
        fromDate = time.mktime(datetime.datetime.strptime(self.settings['from'],"%Y-%m-%d").timetuple())
        print fromDate
        toDate = time.mktime(datetime.datetime.strptime(self.settings['to'],"%Y-%m-%d").timetuple())
        print toDate
        dbParamsM = self.dbConfiguration
        try:
            dbS = psycopg2.connect(database=dbParamsM['db'], user=dbParamsM["username"], host=dbParamsM["host"], port=dbParamsM["port"])
        except psycopg2.Error as e:
            print e
            return False
            pass
        print "Conectado a la DB"
        cur = dbS.cursor()
        try:
            cur.execute("""SELECT sdo.NAME ,
            count(wo.WORKORDERID),
            count(case when wos.ISOVERDUE='1' then 1 else null end),
            count(case when wos.is_fr_overdue='1' then 1 else null end) ,
            case when count(wo.workorderid) > 0 then count(case when (wos.is_fr_overdue='1') THEN 1 ELSE NULL END) *100 / count(wo.workorderid) else null end ,
            case when count(wo.workorderid) > 0 then count(case when (wos.ISOVERDUE='1') THEN 1 ELSE NULL END) *100 / count(wo.workorderid) else null end  FROM WorkOrder wo
            LEFT JOIN SiteDefinition siteDef ON wo.SITEID=siteDef.SITEID
            LEFT JOIN SDOrganization sdo ON siteDef.SITEID=sdo.ORG_ID
            LEFT JOIN WorkOrderStates wos ON wo.WORKORDERID=wos.WORKORDERID
            Left join statusdefinition std on wos.statusid=std.statusid  where wo.CREATEDTIME >=%s AND wo.CREATEDTIME <=%s group by sdo.NAME
            order by 1""", (fromDate, toDate))
        except psycopg2.Error as e:
            print "Error"
            print e
        rows = cur.fetchall()
        print rows


def main():
    report = KPI()
    report.obtainDataFromServiceDesk()

if __name__ == '__main__':
	main()
