#!/usr/bin/env python

import xlrd, os

def totalComplaints():
    for file in os.listdir('files'):
        if not (file.endswith(".xls") or file.endswith("xlsx")):
            continue
        # Open the workbook
        workbook = xlrd.open_workbook("files/" + file)
        
        # Print the workbook name
        print file  + ":"
        
        if('mydtxt' in file.lower()):
            myDtxtCounter(workbook)
            continue

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

            # The starting row since the title rows mean nothing
            row = 2
            
            while(row < worksheet.nrows):    
                if(worksheet.cell(row, 0).value == ""):
                    row += 1
                    continue
                
                # Then check to see if it is facebook or website
                # The current cell
                cellValue = worksheet.cell(row,2).value
                if('facebook' in (cellValue.lower())):
                    facebookComments+=1
                elif('website' in (cellValue.lower())):
                    websiteComments+=1
                row+=1

        # Once the entire workbook has been searched, print out the values of the facebook comments and website comments
        print "\tFacebook Complaints: {}".format(facebookComments)
        print "\tWebsite Complaints: {}".format(websiteComments)
        print ""
    
    
    
def lateNightComplaints():
    for file in os.listdir('.'):
        if not file.endswith(".xls"):
            continue
        # Open the workbook
        workbook = xlrd.open_workbook(file)

        # Print the workbook name
        print file
        
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

            # The starting row since the title rows mean nothing
            row = 2
            
            # The flag to tell if the current row is a late night comment
            lateNight = 0
            
            while(row < worksheet.nrows):
                # First check to see if it is a late night comment
                timeCheck = worksheet.cell(row, 1).value
                
                # Once we have the text in the cell split it by the space between the date and the time
                dateTime = timeCheck.split(' ')
                
                # Check if there is a time 
                if len(dateTime) > 1:
                    print(dateTime[1])
                    # If there is a colon in the time
                    #if(':' in (dateTime[1]):
                       
                
                # Then check to see if it is facebook or website
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

def myDtxtCounter(workbook):
    # Open the first sheet
        worksheet = workbook.sheet_by_index(0)

        # Once the entire workbook has been searched, print out the values of the facebook comments and website comments
        print "\tmyDtxt Complaints: {}".format(worksheet.nrows - 2)
        print ""
    
    
def main():
    totalComplaints()

    
main()