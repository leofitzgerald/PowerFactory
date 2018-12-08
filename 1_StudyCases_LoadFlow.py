#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Sat Dec  8 15:39:23 2018

@author: lex
"""


import powerfactory
app = powerfactory.GetApplication()
# get Load Flow object
ldf = app.GetFromStudyCase('ComLdf')

# Get Study Case Folder
aFolder = app.GetProjectFolder('study')
app.PrintPlain(aFolder) # returns 'Study Cases' folder object

# Get contents of Study Cases folder
aCases = aFolder.GetContents('*.IntCase')
aCases = aCases[0]
# Print study cases available
app.PrintPlain(aCases) # returns 'Case 1 ...' 'Case 2....'

# Loop through each Case Study & Activate it, run load flow
for Case in aCases:
    # get all the Lines in model
    lines = app.GetCalcRelevantObjects("*.ElmLne")
    Case.Activate
    ldf.Execute()
    
    # return the name of all lines
    for line in lines:
        #app.PrintPlain(line) # returns Line OBJECT by name eg 'Line 1'
        line_name = line.loc_name # returns Line name as String eg 'Line 1'
        line_loading = line.GetAttribute('c:loading')
        line_length = line.GetAttribute("b:dline")
        
        # print line name & loading
        app.PrintPlain("{}: {} %".format(line_name,round(line_loading,3)))
        app.PrintPlain("{}: Length: {} km".format(line_name,line_length))
        
        # print flexible Line Data
        app.PrintPlain(app.OutputFlexibleData(lines))

        
        
        
        
        
        
        
        
        
        
        




