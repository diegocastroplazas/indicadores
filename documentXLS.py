#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlsxwriter

class Document:

    def __init__(self, filename):

        self.filename = filename
        self.workbook = xlsxwriter.Workbook(self.filename)
        self.worksheet = self.workbook.add_worksheet()
        self.createColumns()

    def createColumns(self, columnNames, row):

        format = self.workbook.add_format({'bold': True})
        currentWorksheet = self.worksheet
        row = 0
        column = 0
        i = 0
        columnNames = ['Empresa', 'TotalTickets','Vencidos por resolucion','Vencidos por respuesta','Porct Vencidos por respuesta','Porct Vencidos por resolucion']

        for col in columnNames:
            currentWorksheet.write_string(row, column, str(columnNames[i]), format)
            i = i + 1
            column = column + 1

    def fillData(self, data, row, column):

        print "Data "
        print data
        currentWorksheet = self.worksheet

        for item in data:
            print (item)
            c = column
            for i in item:
                currentWorksheet.write_string(row, c, str(i).decode('utf-8'))
                c = c + 1
            row = row + 1
        return row


    def endEdition(self):
        self.workbook.close()
