#!/usr/bin/env python
# -*- coding: utf-8 -*-

from sdp import SDP
from documentXLS import Document

conn = SDP(host='127.0.0.1', user='postgres', passw='', port='65432', fromDate='2017-03-01', toDate='2017-03-31')
result = conn.obtainResume()
excel = Document('archivo.xls')
lastrow = excel.fillData(result,1,0)
print "Ultima fila: " + str(lastrow)
result = conn.obtainRequestsList()
excel.fillData(result,lastrow + 1,0)
excel.endEdition()
