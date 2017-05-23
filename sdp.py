#!/usr/bin/env python
# -*- coding: utf-8 -*-

import psycopg2
import time
import datetime

class SDP:

    def __init__(self, host, user, passw, port, fromDate, toDate):

        self.host = host
        self.user = user
        self.passw = passw
        self.port = port
        self.database = 'servicedesk'
        self.fromDate = fromDate
        self.toDate = toDate

        print "Data inicializada"
        if self.createConn():
            print "Data conectada"
        else:
            print "No se ha conectado"

    def createConn(self):

        print "voy a crear conexion a la DB"
        try:
            self.dbS = psycopg2.connect(database=self.database, user=self.user, host=self.host, port=self.port)
        except psycopg2.Error as e:
            raise ErrorConn("No se puede conectar a la Base de Datos: " + e)
        return True

    def obtainResume(self):

        print "Voy a obtener los datos"
        self.fromDate = int(time.mktime(datetime.datetime.strptime(self.fromDate,"%Y-%m-%d").timetuple())) * 1000
        print self.fromDate
        self.toDate = int(time.mktime(datetime.datetime.strptime(self.toDate,"%Y-%m-%d").timetuple())) * 1000
        print self.toDate
        cur = self.dbS.cursor()
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
            order by 1""", (self.fromDate, self.toDate))
        except psycopg2.Error as e:
            print "Error"
            raise ErrObtainKPI("No es posible obtener los datos: " + e)
        rows = cur.fetchall()
        print rows
        return rows

    def obtainRequestsList(self):

        print "Voy a obtener datos de todos los tickets"
        cur = self.dbS.cursor()
        print self.fromDate
        print self.toDate
        try:
            cur.execute("""SELECT sdo.NAME,
            wo.WORKORDERID,
            wo.TITLE,
            std.STATUSNAME,
            to_char(to_timestamp(wo.CREATEDTIME/1000),'DD-MM-YYYY HH24:MI:SS'),
            to_char(to_timestamp(wo.RESPONDEDTIME/1000),'DD-MM-YYYY HH24:MI:SS'),
            to_char(to_timestamp(wo.RESOLVEDTIME/1000),'DD-MM-YYYY HH24:MI:SS'),
            wos.IS_FR_OVERDUE,
            wos.ISOVERDUE FROM WorkOrder wo
            LEFT JOIN SiteDefinition siteDef ON wo.SITEID=siteDef.SITEID
            LEFT JOIN SDOrganization sdo ON siteDef.SITEID=sdo.ORG_ID
            LEFT JOIN WorkOrderStates wos ON wo.WORKORDERID=wos.WORKORDERID
            LEFT JOIN StatusDefinition std ON wos.STATUSID=std.STATUSID
            where wo.CREATEDTIME >=%s AND wo.CREATEDTIME <=%s  ORDER BY 1 NULLS FIRST""", (self.fromDate, self.toDate))
        except psycopg2.Error as e:
            print "Error"
            print e
        rows = cur.fetchall()
        return rows
