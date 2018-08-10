#!/usr/bin/env python
import sys
import time
import os
from entity_reader import entity_reader

def main(argv):
    
    Myfilepath='/Users/itadmin/Documents/IAEA/Dropbox_Mirror/Long Reports/'
    getdoi=1
    debug=0 ## Set to 1 for debug mode - does not write data
    fileslist=[]

    ###Get list of files to open
    for root, dirs, files in os.walk(Myfilepath):
        for file_ in files:
            if file_.endswith(".docx"):
                ###print(os.path.join(root, file_))
                fileslist.append(os.path.join(root, file_))

    Myfilepath2='/Users/itadmin/Documents/IAEA/Archives/'
    for root, dirs, files in os.walk(Myfilepath2):
        for file_ in files:
            if file_.endswith(".docx"):
                ###print(os.path.join(root, file_))
                fileslist.append(os.path.join(root, file_))
    
    #fileslist=['/Users/itadmin/Documents/IAEA/Long Reports/ME/KCL OS Updates EA - Feb 18.docx']
    #fileslist=['/Users/itadmin/Documents/IAEA/Dropbox_Mirror/Long Reports/Middle East & North Africa/2017/Syria Intl coop 2007-17 - Sep 17.docx']
    #fileslist=['/Users/itadmin/Documents/IAEA/Archives/MENA/KCL MENA Open Source Nuclear Profiles - Three/Turkey Nuclear Profile - Nov 10.docx']
    fileslist=['/Users/itadmin/Documents/IAEA/Dropbox_Mirror/Long Reports/Middle East & North Africa/2017/Egypt Nuclear Profile V7 - Mar 17.docx']
    ###For that list get args for entity reader            
    for file_ in fileslist:
        Myfilename=file_.split('/').pop()
        report_name=Myfilename.split('-')[0].strip()
        Month=Myfilename.split('-').pop().strip().split(' ')[0]
        Year=Myfilename.split('-').pop().strip().split(' ').pop().split('.')[0]
        print file_
        print "Report Name=", report_name, 'Month=',Month, 'Year=',Year 
        entity_reader(file_, report_name, Month, Year, getdoi, debug)

        for i in xrange(0,1):
            time.sleep(1)
            sys.stdout.write('...\r')
            sys.stdout.flush()
        print '\n'    


if __name__=="__main__":
    main(sys.argv[1:])
