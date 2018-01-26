# -*- coding: utf-8 -*-
"""
Created on Thu Jan 25 19:14:58 2018

@author: slwis
"""

import Pep_data_extraction as ext
import numpy


crop_info, jdates, gc1rep, gc2rep, gc3rep = ext.parse_file("PEP12014040.xlsx")

#print(jdates,crop_info)



def write_datefile(crop_info, jdates):
    
    with open("pep_dates.txt", "w") as g:
        
        #line = "cropname: {}".format(crop_info["cropName"])
        g.write(crop_info["cropName"]) 
        g.write("\t")
        g.write(crop_info["year"])
        g.write("\t")
        
        #list comprehension
        jdates_str = [str(x) for x in jdates]
        
        data = "\t".join(jdates_str)
        
        g.write(data)

write_datefile(crop_info, jdates)

gcreps = [gc1rep, gc2rep, gc3rep]
print(gcreps)
GCjd2 = [] # trying to average the values of the 3 reps for each measurement point
for x in gcreps:
    for n in x:
        print(n)
        GCjd2.append(x[n])
        print(GCjd2)
        aveGC2 = numpy.mean(GCjd2)
        print(x[n-1],aveGC2)


def write_detailsfile(crop_info, gc1rep, gc2rep, gc3rep):
    with open ("pep_details.txt","w") as g:
        g.write(crop_info["cropName"])
        g.write("\t")
        g.write(crop_info["cropVar"])
        g.write("\t")        
        g.write(str(crop_info["NoR"]))
        g.write("\t")        
        g.write(crop_info["year"])
        g.write("\t")        
        g.write(crop_info["pdate"])
        g.write("\t")        
        g.write(crop_info["emdate"])
        g.write("\t")        
        g.write(crop_info["met_ref"])
        g.write("\t")        
        g.write(str(crop_info["stem_den1"])) #need to limit to 0dp
        g.write("\t")        
        g.write(str(crop_info["stem_den2"]))  #need to limit to 0dp
        g.write("\t")        
        g.write(str(crop_info["detlev"]))
        g.write("\t")       
        g.write(str(crop_info["seed_mass"])) #need to restrict to 1dp
        g.write("\t")       
        g.write(str(crop_info["seed_spacing"])) #need to restrict to 0dp
        g.write("\t")        
        g.write(str(crop_info["row_width"]))
        g.write("\t")        
        g.write(str(crop_info["planting_density"]))
        g.write("\t")        
        g.write(str(crop_info["appN"]))
        g.write("\t")
        
        jdates_str = [str(x) for x in jdates]
        
        data = "\t".join(jdates_str)
        
        g.write(data)
#        crop_info_str = [str(x) for x in crop_info] # this just produces titles, the keys(?) of the dictionary
#        cdata = "\t".join(crop_info_str)
#        g.write("\n")
#        g.write(cdata)
#        for date in jdates_str: #for changing the orientation of the dates in the string
#            g.write(date + "\n")
#            
#

write_detailsfile(crop_info, gc1rep, gc2rep, gc3rep)

#
#parse_file("PEP12014040.xlsx")
#
#all_crops=[] #trying to count the number of crops but crop is currently empty
#all_crops.append(crop) #also not sure where it needs to be located in script
#noc=len(all_crops)
#
#TEMP_FILE = 'template.txt'
#OUTPUT_FILE = 'all_output.txt'
#
#def append_to_temp_file(crop):
#    """
#    {
#      'name': value (str)  
#      'variety' : value (str)
#      'number of days (nod)': value (int)
#    }
#    """
#    template = '{name}\t{variety}\t{nod}'
#    with open(TEMP_FILE, 'a') as temp:
#        temp.write(template.format(**crop))
#
#def finalise_output(noc):
#    with open(OUTPUT_FILE, 'w') as output_file:
#        output_file.write('{}\n:\n'.format(noc))
#        with open(TEMP_FILE, 'r') as in_temp:
#            output_file.write(in_temp.read())
#
#def main():
#    for file_name in os.listdir('.'):
#        if file_name.endswith('.xlsx'):
#            print('parsing', file_name)
#            crop = parse_file(file_name)
#            append_to_temp_file(crop)
#    finalise_output()
#
#
#if __name__ == '__main__':
#    main()