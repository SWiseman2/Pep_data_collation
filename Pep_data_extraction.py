# -*- coding: utf-8 -*-

import os

import openpyxl
from openpyxl import load_workbook

    #dictionary for looking up determinacy level
VarietyDeterminacy = {"Accord": 1,
                      "Agria":	3,
                      "Ambo":	3,
                      "Amora":	3,
                      "Annabelle":	1,
                      "Anya":	1,
                      "Arsenal":3,
                      "Asterix":4,
                      "Atlantic":	2,
                      "Brooke":	4,
                      "Cabaret":	3,
                      "Caesar":	3,
                      "Cara":4,
                      "Carlingford":	2,
                      "Charlotte":	2,
                      "Colmo":1,
                      "Cosmos":	3,
                      "Courage":	2,
                      "Cultra":	3,
                      "Daisy":	3,
                      "Desirée":	3,
                      "Dundrod":	2,
                      "Estima":	1,
                      "Fambo":	3,
                      "Fianna":	3,
                      "Harmony":	2,
                      "Hermes":	3,
                      "Innovator":	1,
                      "Infinity": 'na',
                      "Juliette":	2,
                      "Kerr’s Pink":	3,
                      "Kestrel":	2,
                      "King Edward":	3,
                      "Lady Balfour":	4,
                      "Lady Christl":	3,
                      "Lady Claire":	2,
                      "Lady Rosetta"	:2,
                      "Marfona":	2,
                      "Maris Bard":	1,
                      "Maris Peer":	2,
                      "Maris Piper":	3,
                      "Maritiema":	2,
                      "Markies":	4,
                      "Melody":	2,
                      "Minerva":	1,
                      "Morene":	3,
                      "Mozart":	2,
                      "Nadine":	2,
                      "Navan":	3,
                      "Nicola":	2,
                      "Olympus": 'na',
                      "Orla":	2,
                      "Osprey":	2,
                      "Pentland Dell":	3,
                      "Pentland Javelin":	2,
                      "Picasso":	3,
                      "Premiere":	1,
                      "Record":	3,
                      "Rembrandt":	2,
                      "Rocket":	1,
                      "Romano":	2,
                      "Rooster":	3,
                      "Russet Burbank":	3,
                      "Sante":	3,
                      "Sassy":	3,
                      "Saturna":	3,
                      "Saxon":	2,
                      "Shannon":	2,
                      "Shelford":	3,
                      "Shepody":	2,
                      "Slaney":	3,
                      "Stemster":	3,
                      "Vales Emerald":	1,
                      "Vales Everest":	4,
                      "Vales Sovereign":	4,
                      "Valor":	3,
                      "Victoria":	3,
                      "Vivaldi":	2,
                      "VR808":	2,
                      "Wilja":	2,
                      "Winston":	1  }


def parse_file(file_name):#defining where the data will be extracted from in 
              #original worksheets, this will actually be carried out later

    with open(file_name, "rb") as my_file:# Read in worksheet from workbook in one folder
             workbook=load_workbook(my_file, data_only=True) #data_only gets 
                                #openpyxl to evaluate the formaul in the cells
             wsheet=workbook.active
     
    lowerBSeed = wsheet["E67"].value #read in lower bound of seed size
    upperBSeed = wsheet["G67"].value # read in upper bound of seed size
    rangeSeed = "%s_%s" % (lowerBSeed,upperBSeed) #formatting seed size ranges
    aveSeed = (lowerBSeed + upperBSeed)/2
                      
    seedmass = 50000/wsheet["E68"].value     #calculating average seed mass
    
    #stem density calculations
    eStemR1 = wsheet["D340"].value #number of stems for first rep (first harvest)
    eStemR2 = wsheet["D341"].value #number of stems for second rep (first harvest)
    eStemR3 = wsheet["D342"].value #number of stems for third rep (first harvest)
    ePlantR1 = wsheet["C340"].value #number of plants in sample (first harvest)
    ePlantR2 = wsheet["C341"].value #number of plants in sample (first harvest)
    ePlantR3 = wsheet["C342"].value #number of plants in sample (first harvest)
    density = wsheet["E36"].value #plant density
    earlystems = ((eStemR1 + eStemR2 + eStemR3)/(ePlantR1 + ePlantR2 + \
                  ePlantR3)) * density #stems per hectare at first harvest
    #print('earlystems',earlystems)
    
    lStemR1 = wsheet["D340"].value #number of stems for first rep (second harvest)
    lStemR2 = wsheet["D341"].value #number of stems for second rep (second harvest)
    lStemR3 = wsheet["D342"].value #number of stems for third rep (second harvest)
    lPlantR1 = wsheet["C340"].value #number of plants in sample (second harvest)
    lPlantR2 = wsheet["C341"].value #number of plants in sample (second harvest)
    lPlantR3 = wsheet["C342"].value #number of plants in sample (second harvest)
    density = wsheet["E36"].value 
    latestems = ((lStemR1 + lStemR2 + lStemR3)/(lPlantR1 + lPlantR2 + \
                 lPlantR3)) * density #stems per hectare at second harvest
    #print('latestems',latestems)
 
    if wsheet["F108"].value is None: #reading in value of total intended fertilizer 
        nfertilizer = "NA"
    else:
        nfertilizer = wsheet["F108"].value
    #print(nfertilizer)
    

    crops = wsheet["E31"].value
    cropDetLev = VarietyDeterminacy[crops]
    #print(crops,cropDetLev)

    
    jdates = []
    crop_dates = {"dates":jdates}
    gc1rep = []
    gc2rep = []
    gc3rep = []
    
    
    cropyear = (wsheet["E143"].value).strftime("%Y")
    
    jdcells = ["E143","E191","F191","G191","H191","I191","J191","K191",\
               "L191","M191","D200","E200","F200","G200","H200","I200",\
               "J200","K200","L200","M200"]
    NoV = 0
    for cell in jdcells:
        try: # this try stops python from trying to read in values which aren't there in the datasheet
            jdcel = wsheet[cell].value.strftime("%j")
        except AttributeError:
            continue  # sonetimes the cells are emoty and this is fine
        jdcell = int(jdcel) #converting to an interger
        #print("jdcell",jdcel)
        if jdcell > 0:
            NoV +=1 
            jdates.append(jdcell)
            #print(NoV, jdcell, jdates)
        else:
            pass
        #print(NoV)
    
    # Dan's dirty hack
    # row1 = ['{}192'.format(chr(i)) for i in range(68, 78)] 
    # row2 = ['{}202'.format(chr(i)) for i in range(68, 78)] 
    # gc1cells = row1 + row2
    gc1cells = ["D192","E192","F192","G192","H192","I192","J192","K192",\
                "L192","M192","D201","E201","F201","G201","H201","I201",\
                "J201","K201","L201","M201"]
    NoGC1 = 0       
    for cell in gc1cells:
        gc1cel = wsheet[cell].value
        if gc1cel is None:
            continue
        gc1cell = int(gc1cel) 
        #print("gc1cell",gc1cel)
        if gc1cell >= 0:
            NoGC1 +=1
            gc1rep.append(gc1cell)
            print(NoGC1, gc1cell, gc1rep)
        
    gc2cells = ["D193","E193","F193","G193","H193","I193","J193","K193",\
                "L193","M193","D202","E202","F202","G202","H202","I202",\
                "J202","K202","L202","M202"]
    NoGC2 = 0       
    for cell in gc2cells:
        gc2cel = wsheet[cell].value
        if gc2cel is None:
            continue
        gc2cell = int(gc2cel) 
        #print("gc2cell",gc2cel)
        if gc2cell >= 0:
            NoGC2 +=1
            gc2rep.append(gc2cell)
            #print(NoGC2, gc2cell, gc2rep)

    gc3cells = ["D194","E194","F194","G194","H194","I194","J194","K194",\
                "L194","M194","D203","E203","F203","G203","H203","I203",\
                "J203","K203","L203","M203"]
    NoGC3 = 0       
    for cell in gc3cells:
        gc3cel = wsheet[cell].value
        if gc3cel is None:
            continue
        gc3cell = int(gc3cel)
        #print("gc3cell",gc3cel)
        if gc3cell >= 0:
            NoGC3 +=1
            gc3rep.append(gc3cell)
            print(NoGC3, gc3cell, gc3rep)
    
    REPs = 0 #counting number of reps
    print(len(gc1rep))
    if len(gc1rep) != 0: #if therere ground cover values in rep 1
          REPs += 1
          print("1 rep",REPs)
          if len(gc2rep) != 0:
              REPs += 1
              if len(gc3rep) != 0:
                  REPs += 1
   # print("REPs",REPs)

    crop_info = {"cropName": wsheet["D1"].value, #crop reference name/number
                  "cropVar": wsheet["E31"].value, #potato variety 
                  "NoD": NoV,     #number of days in each indiv crop
                  "NoR": REPs,      #number of reps 
                      #(sequential, if entered rep1+, rep2-, rep3+ would count as 1) - not sure what I meant here....
                  "year": wsheet["J33"].value.strftime("%Y") #reads in year from pdate 
                 
                  "pdate": wsheet["J33"].value.strftime("%j"), #reads planting 
                                            #date and converts to Julian Date
                  "emdate": wsheet["E143"].value.strftime("%j"),#ditto emergence 
                  "met_ref": wsheet["F59"].value, #reads met data reference
                  "stem_den1": earlystems, #including early stem density
                  "stem_den2": latestems, #inc late stem density
                  "detlev": cropDetLev,# looks up detlev from VarietyDeterminacy
                  "seed_mass": seedmass, #including seed mass in the dictionary
                  "seed_grade": rangeSeed, #inc raw seed garde
                  "ave_seed_grade": aveSeed, #inc middle point of seed grade
                  "seed_spacing": wsheet["E37"].value,
                  "row_width": wsheet["E25"].value,#reads in value of row width 
                  "planting_density": wsheet["E36"].value,
                  "appN": nfertilizer #read in applied nitrogen
                  } 
    
       
    
    return [crop_info, jdates, gc1rep, gc2rep, gc3rep]

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