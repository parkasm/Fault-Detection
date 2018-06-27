##STEPS TO COMPLETE BEFORE RUNNING THIS SCRIPT
# STEP 1 : HAVE THE EXCEL DATA READY AND UPDATED WITH CURRENT AND VOLTAGE NUMBERS
# STEP 2 : HAVE THE CURRENT MAGNITUDES AND THE THRESHOLD CURRENT VALUES. UPDATE THE FAULT_TYPE SCRIPT WITH THE THRESHOLD CURRENT VALUE 
# STEP 3 : RUN THIS SCRIPT. PASS THOSE CURRENT MAGNITUDES IN LINE 14 TO FAULT TYPE FUNCTION. PLAN IS TO RUN THIS SCRIPT DIRECTLY WITHOUT PASSING IN ANY ARGUMENTS


from fault_type import fault_type # this is importing the fault type function from the fault type file
import xlrd
import panda
import numpy as np
import cmath
import math
import sys
f = fault_type(279.6,246.9,145.2)
## Different values of f corresponds to different kind of faults. There is a dictionary created in fault_type.py file to look into details of the values of f
#f = 10

# This is the core function which outputs the section containing the fault. The input for this function are the Modified Reactance for a single section of distribution line and the apparent reactance from the kind of fault.  |
# The modified reactance equations can be found in the Ratan Das Paper. If the Modified Reactance for a section is less than the Apparent Reactance, which means the fault is not in that section and look for the fault in the next sections. 

def Section_Detection_Function(ApparentReactance):

    Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault = ((Pos_Seq_Inductance_Henry_per_km * 2 * cmath.pi * 60 * Section_Length) + (((Zero_Seq_Inductance_Henry_per_km - Pos_Seq_Inductance_Henry_per_km) * 2 * cmath.pi * 60 * Section_Length)/3))
    Modified_Reactance_Per_Section_Other_Faults = (Pos_Seq_Inductance_Henry_per_km * 2 * cmath.pi * 60 * Section_Length)
    print('Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault is %f' % Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault)
    print('Modified_Reactance_Per_Section_Other_Faults is %f' % Modified_Reactance_Per_Section_Other_Faults)
    x = 1
    y = 1

    ## The parameters for fault detection is same for the Single Line to Ground Faults and same for the rest of the faults. That is why there are blocks of if-else statements to check for the value of f and based on that calculations are done
    if f == 1 or f == 2 or f == 3:
        for j in range(1,Total_Number_of_Line_Sections + 1):
            if (Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault < ApparentReactance):
                print('Fault is beyond this section')
                #If the Modified Reactance for a section is less than the Apparent Reactance, which means the fault is not in that section and look for the fault in the next sections. Modified Reactance changes when moved to next section. 
                Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault = Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault + Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault
                x = x + 1
            else:
                print('Fault is in Section %d' % x )
                        
    if f == 4 or f == 5 or f == 6  or f == 7 or f == 8 or f == 9 or f == 10:
        for k in range(1,Total_Number_of_Line_Sections + 1):
            if (Modified_Reactance_Per_Section_Other_Faults < ApparentReactance):
                print('Fault is beyond this section' )
                Modified_Reactance_Per_Section_Other_Faults = Modified_Reactance_Per_Section_Other_Faults + Modified_Reactance_Per_Section_Other_Faults
                y = y + 1
            else:
                print('Fault is in Section %d' % y )

    return  True


## Distribution Line Parameters. Coming from Electrical Power Distribution Handbook
Pos_Seq_Resistace_ohm_per_km = 0.188272
Zero_Seq_Resistace_ohm_per_km = 0.5540
## To Calculate Reactance just multiply inductance with 2 * Pi * 60 or ~377 as well by the section length
Pos_Seq_Inductance_Henry_per_km = 1.122e-3 
Zero_Seq_Inductance_Henry_per_km = 3.55e-3

## Default Capacitance Sequence values in simulink were used
Pos_Seq_Capacitance_Farad_per_km = 10.74e-9
Zero_Seq_Capacitance_Farad_per_km = 5.75e-9

## Section Length of the distribution Network. Used 6 km and 20 Km for the two simulink models created
Section_Length = int(input('Enter the Section length in kilometer:'))
Total_Number_of_Line_Sections = 4 ## THIS NUMBER IS BASED ON Simulink_Model.slx file

##Loading the Excel Workbooks containing all the current, voltage and sequence component information
workbook = xlrd.open_workbook('testfile_new.xls')
workbook_2 = xlrd.open_workbook('Sequence_Components.xls')

##Assigning the sheets of the workbook to the respective parameters. 
##IT IS REALLY IMPORTANT TO HAVE THE DATA IN THE SEQUENCE MENTIONED BELOW OTHERWISE ALL THE CALCULATIONS WILL BE WRONG. ALTERNATIVELY, IF NEED TO MAKE CHANGES IN THE SEQUENCE THEN ALSO MAKE CHANGES IN THE CALCULATION DONE FOR EACH FAULT

# In workbook, The First sheet contains the voltage in complex form
Voltage_SpreadSheet = workbook.sheet_by_index(0)

# In workbook, The Second sheet contains the current in complex form
Current_SpreadSheet = workbook.sheet_by_index(1)

# In workbook, The Third sheet contains only the real components of the voltage 
Voltage_SpreadSheet_Real_Part = workbook.sheet_by_index(2)

# In workbook, The Fourt sheet contains only the imaginary component of the voltage
Voltage_SpreadSheet_Imag_Part = workbook.sheet_by_index(3)

# In workbook, The fifth sheet contains only the real components of the current
Current_SpreadSheet_Real_Part = workbook.sheet_by_index(4)

# In workbook, The Sixth sheet contains only the imaginary components of the current
Current_SpreadSheet_Imag_Part = workbook.sheet_by_index(5)

#In workbook_2, the first sheet contains the magnitude of the positive sequence of voltage 
Voltage_Sequence_Magnitude = workbook_2.sheet_by_index(0)

#In workbook_2, the second sheet contains the phase/angle in radians of the positive sequence of voltage 
Voltage_Sequence_Angle = workbook_2.sheet_by_index(1)

#In workbook_2, the third sheet contains the maginitude of the positive sequence of current
Current_Sequence_Magnitude = workbook_2.sheet_by_index(2)

#In workbook_2, the fourth sheet contains the phase/angle in radians of the positive sequence of current
Current_Sequence_Angle = workbook_2.sheet_by_index(3)     

""" print(Voltage_SpreadSheet.nrows)
print(Current_SpreadSheet.nrows)
print(Voltage_SpreadSheet.cell_value(0,0))
val1= Voltage_SpreadSheet_Real_Part.cell_value(2,1)
print(val1) """

##BASED ON THE FAULT TYPE, THE APPARENT REACTANCE IS BEING CALCULATED AND THEN THAT IS INPUT TO THE SECTION_DETECTION_FUNCTION.
## The Equations for the Apparent Reactance can be found in the Ratan DAS Paper. Those are implemented below in the conditional loops which are based on the values of 'f' aka the type of fault.

# Small check to make sure that there are equal number of current and voltage phasors in the excel sheets otherwise there is something wrong
if(Voltage_SpreadSheet.nrows == Current_SpreadSheet.nrows):
    #Loop through all the iterations of the voltage/current phasors to get the apparent reactance for each of them
    for i in range(0,Voltage_SpreadSheet.nrows):
            if f == 0 :
                print('There is no fault')
            elif f == 1 :
                Apparent_Impedance_AG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0)))/(complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0))))
                Apparent_Reactance_AG = abs(Apparent_Impedance_AG.imag)
                print('Apparent_Impedance_AG: %f+i%f' % (Apparent_Impedance_AG.real,Apparent_Impedance_AG.imag))
                print('Apparent_Reactance_AG: %f' % Apparent_Reactance_AG)
                Section_Detection_Function(Apparent_Reactance_AG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 2 :
                Apparent_Impedance_BG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1)))/(complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1))))
                Apparent_Reactance_BG = abs(Apparent_Impedance_BG.imag)
                print('Apparent_Impedance_BG: %f+i%f' % (Apparent_Impedance_BG.real,Apparent_Impedance_BG.imag))
                print('Apparent_Reactance_BG: %f' % Apparent_Reactance_BG)
                Section_Detection_Function(Apparent_Reactance_BG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 3 :
                Apparent_Impedance_CG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2)))/(complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2))))
                Apparent_Reactance_CG = abs(Apparent_Impedance_CG.imag)
                print('Apparent_Impedance_CG: %f+i%f' % (Apparent_Impedance_CG.real,Apparent_Impedance_CG.imag))
                print('Apparent_Reactance_CG: %f' % Apparent_Reactance_CG)
                Section_Detection_Function(Apparent_Reactance_CG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 4 :
                Numerator_ABG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1))))
                Denominator_ABG = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1))))
                Apparent_Impedance_ABG = Numerator_ABG / Denominator_ABG
                Apparent_Reactance_ABG = abs(Apparent_Impedance_ABG.imag)
                print('Apparent_Impedance_ABG: %f+i%f' % (Apparent_Impedance_ABG.real,Apparent_Impedance_ABG.imag))
                print('Apparent_Reactance_ABG: %f' % Apparent_Reactance_ABG)
                Section_Detection_Function(Apparent_Reactance_ABG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 5 :
                Numerator_BCG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2))))
                Denominator_BCG = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2))))
                Apparent_Impedance_BCG = Numerator_BCG / Denominator_BCG
                Apparent_Reactance_BCG = abs(Apparent_Impedance_BCG.imag)
                print('Apparent_Impedance_BCG: %f+i%f' % (Apparent_Impedance_BCG.real,Apparent_Impedance_BCG.imag))
                print('Apparent_Reactance_BCG: %f' % Apparent_Reactance_BCG)
                Section_Detection_Function(Apparent_Reactance_BCG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 6 :
                Numerator_CAG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0))))
                Denominator_CAG = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0))))
                Apparent_Impedance_CAG = Numerator_CAG / Denominator_CAG
                Apparent_Reactance_CAG = abs(Apparent_Impedance_CAG.imag)
                print('Apparent_Impedance_CAG: %f+i%f' % (Apparent_Impedance_CAG.real,Apparent_Impedance_CAG.imag))
                print('Apparent_Reactance_CAG: %f' % Apparent_Reactance_CAG)
                Section_Detection_Function(Apparent_Reactance_CAG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 7 :
                Numerator_AB = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1))))
                Denominator_AB = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1))))
                Apparent_Impedance_AB = Numerator_AB / Denominator_AB
                Apparent_Reactance_AB = abs(Apparent_Impedance_AB.imag)
                print('Apparent_Impedance_AB: %f+i%f' % (Apparent_Impedance_AB.real,Apparent_Impedance_AB.imag))
                print('Apparent_Reactance_AB: %f' % Apparent_Reactance_AB)
                Section_Detection_Function(Apparent_Reactance_AB) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 8 :
                Numerator_BC = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2))))
                Denominator_BC = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2))))
                Apparent_Impedance_BC = Numerator_BC / Denominator_BC
                Apparent_Reactance_BC = abs(Apparent_Impedance_BC.imag)
                print('Apparent_Impedance_BC: %f+i%f' % (Apparent_Impedance_BC.real,Apparent_Reactance_BC.imag))
                print('Apparent_Reactance_BC: %f' % Apparent_Reactance_BC) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 9 :
                Numerator_CA = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0))))
                Denominator_CA = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0))))
                Apparent_Impedance_CA = Numerator_CA / Denominator_CA
                Apparent_Reactance_CA = abs(Apparent_Impedance_CA.imag)
                print('Apparent_Impedance_CA: %f+i%f ' % (Apparent_Impedance_CA.real,Apparent_Impedance_CA.imag))
                print('Apparent_Reactance_CA: %f' % Apparent_Reactance_CA)
                Section_Detection_Function(Apparent_Reactance_CA) # Calling the Section_Detection Function here with Apparent Reactance as the argument
            elif f == 10 :
                Numerator_ABCG = cmath.rect((Voltage_Sequence_Magnitude.cell_value(i,0)),(Voltage_Sequence_Angle.cell_value(i,0)))
                Denominator_ABCG = cmath.rect((Current_Sequence_Magnitude.cell_value(i,0)),(Current_Sequence_Angle.cell_value(i,0)))
                Apparent_Impedance_ABCG = Numerator_ABCG / Denominator_ABCG
                Apparent_Reactance_ABCG = abs(Apparent_Impedance_ABCG.imag)
                print('Apparent_Impedance_ABCG: %f+i%f' % (Apparent_Impedance_ABCG.real,Apparent_Impedance_ABCG.imag))
                print('Apparent_Reactance_ABCG: %f' % Apparent_Reactance_ABCG)
                Section_Detection_Function(Apparent_Reactance_ABCG) # Calling the Section_Detection Function here with Apparent Reactance as the argument
else:
    print('Need equal number of current and voltage phasors. Check the simulation model and make sure there are measurement blocks for all the phasors')

