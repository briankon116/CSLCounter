#!/usr/bin/env python

import xlrd, os

def totalComplaints():
    for file in os.listdir('.'):
        if not file.endswith(".xls"):
            continue
        # Open the workbook
        workbook = xlrd.open_workbook(file)

        # Open the first sheet
        worksheet = workbook.sheet_by_index(0)

        # The counter for the facebook page complaints
        facebookComments = 0

        # The counter for the website comments
        websiteComments = 0

        # Loop over all of the sheets in the workbook
        for i in range (0, len(workbook.sheet_names())):
            # Open the current sheet
            worksheet = workbook.sheet_by_index(i)

            row = 2
            while(row < worksheet.nrows):
                # The current cell
                cellValue = worksheet.cell(row,2).value
                if('facebook' in (cellValue.lower())):
                    facebookComments+=1
                elif('website' in (cellValue.lower())):
                    websiteComments+=1
                row+=1

        # Once the entire workbook has been searched, print out the values of the facebook comments and website comments
        print "Facebook Complaints: {}".format(facebookComments)
        print "Website Complaints: {}".format(websiteComments)
        
def main():
    totalComplaints()

    
main()