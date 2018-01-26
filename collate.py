# -*- coding: utf-8 -*-
"""
Created on Thu Jan 25 19:14:58 2018

@author: slwis
"""

import Pep_data_extraction as ext


crop_info, jdates, gc1rep, gc2rep, gc3rep = ext.parse_file("PEP12014040.xlsx")

#print(jdates,crop_info)



def write_datefile(crop_info, jdates):
    
    with open("pep_dates.txt", "w") as g:
        
        g.write('"pep_dates"\n')
        
        #line = "cropname: {}".format(crop_info["cropName"])
        g.write(crop_info["cropName"]) #+ "\t"
        g.write(crop_info["year"])
        #g.write("\n")  don't need to add an extra line here
        
        #list comprehension
        jdates_str = [str(x) for x in jdates]
        
        data = "\t".join(jdates_str)
        
        g.write(data)
#        
#        for date in jdates_str:
#            g.write(date + "\n")
#            
#        
write_datefile(crop_info, jdates)