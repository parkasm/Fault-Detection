from fault_type import fault_type
import xlrd
import panda
import numpy as np
import cmath
import math
import sys
##f = fault_type(130,140,180)
f = 1


def Section_Detection_Function(ApparentReactance):

    Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault = ((Pos_Seq_Inductance_Henry_per_km * 2 * cmath.pi * 60 * Section_Length) + (((Zero_Seq_Inductance_Henry_per_km - Pos_Seq_Inductance_Henry_per_km) * 2 * cmath.pi * 60 * Section_Length)/3))
    Modified_Reactance_Per_Section_Other_Faults = (Pos_Seq_Inductance_Henry_per_km * 2 * cmath.pi * 60)
    print(Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault)
    x = 1
    if f == 1 or f == 2 or f == 3:
        for j in range(1,Total_Number_of_Line_Sections + 1):
            if (Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault < ApparentReactance):
                print('Fault is beyond this section')
                Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault = Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault + Modified_Reactance_Per_Section_Single_Line_to_Ground_Fault
                x = x + 1
            else:
                print('Fault is in Section %d' % x )
                        
    if f == 4 or f == 5 or f == 6  or f == 7 or f == 8 or f == 9 or f == 10:
        for k in range(1,Total_Number_of_Line_Sections + 1):
            if (Modified_Reactance_Per_Section_Other_Faults < ApparentReactance):
                print('Fault is beyond this section' )
                Modified_Reactance_Per_Section_Other_Faults = Modified_Reactance_Per_Section_Other_Faults + Modified_Reactance_Per_Section_Other_Faults
            else:
                print('Fault is in Section %d' % k )

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

Section_Length = int(input('Enter the Section length in km:'))
Total_Number_of_Line_Sections = 4 ## THIS NUMBER IS BASED ON Simulink_Model.slx file

##Loading the Excel Workbooks
workbook = xlrd.open_workbook('testfile_new.xls')
workbook_2 = xlrd.open_workbook('Sequence_Components.xls')

##Assigning the sheets of the workbook to the respective parameters
Voltage_SpreadSheet = workbook.sheet_by_index(0)
Current_SpreadSheet = workbook.sheet_by_index(1)
Voltage_SpreadSheet_Real_Part = workbook.sheet_by_index(2)
Voltage_SpreadSheet_Imag_Part = workbook.sheet_by_index(3)
Current_SpreadSheet_Real_Part = workbook.sheet_by_index(4)
Current_SpreadSheet_Imag_Part = workbook.sheet_by_index(5)
Voltage_Sequence_Magnitude = workbook_2.sheet_by_index(0)
Voltage_Sequence_Angle = workbook_2.sheet_by_index(1)
Current_Sequence_Magnitude = workbook_2.sheet_by_index(2)
Current_Sequence_Angle = workbook_2.sheet_by_index(3)     

""" print(Voltage_SpreadSheet.nrows)
print(Current_SpreadSheet.nrows)
print(Voltage_SpreadSheet.cell_value(0,0))
val1= Voltage_SpreadSheet_Real_Part.cell_value(2,1)
print(val1) """

##Looping through the rows of Voltage and Current Parameters to get Impedance
if(Voltage_SpreadSheet.nrows == Current_SpreadSheet.nrows):
    for i in range(0,Voltage_SpreadSheet.nrows):
            if f == 0 :
                print('There is no fault')
            elif f == 1 :
                Apparent_Impedance_AG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0)))/(complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0))))
                Apparent_Reactance_AG = abs(Apparent_Impedance_AG.imag)
                print('Apparent_Impedance_AG: %f+i%f' % (Apparent_Impedance_AG.real,Apparent_Impedance_AG.imag))
                print('Apparent_Reactance_AG: %f' % Apparent_Reactance_AG)
                Section_Detection_Function(Apparent_Reactance_AG)
            elif f == 2 :
                Apparent_Impedance_BG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1)))/(complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1))))
                Apparent_Reactance_BG = abs(Apparent_Impedance_BG.imag)
                print('Apparent_Impedance_BG: %f+i%f' % (Apparent_Impedance_BG.real,Apparent_Impedance_BG.imag))
                print('Apparent_Reactance_BG: %f' % Apparent_Reactance_BG)
                Section_Detection_Function(Apparent_Reactance_BG)
            elif f == 3 :
                Apparent_Impedance_CG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2)))/(complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2))))
                Apparent_Reactance_CG = abs(Apparent_Impedance_CG.imag)
                print('Apparent_Impedance_CG: %f+i%f' % (Apparent_Impedance_CG.real,Apparent_Impedance_CG.imag))
                print('Apparent_Reactance_CG: %f' % Apparent_Reactance_CG)
                Section_Detection_Function(Apparent_Reactance_CG)
            elif f == 4 :
                Numerator_ABG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1))))
                Denominator_ABG = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1))))
                Apparent_Impedance_ABG = Numerator_ABG / Denominator_ABG
                Apparent_Reactance_ABG = abs(Apparent_Impedance_ABG.imag)
                print('Apparent_Impedance_ABG: %f+i%f' % (Apparent_Impedance_ABG.real,Apparent_Impedance_ABG.imag))
                print('Apparent_Reactance_ABG: %f' % Apparent_Reactance_ABG)
                Section_Detection_Function(Apparent_Reactance_ABG)
            elif f == 5 :
                Numerator_BCG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2))))
                Denominator_BCG = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2))))
                Apparent_Impedance_BCG = Numerator_BCG / Denominator_BCG
                Apparent_Reactance_BCG = abs(Apparent_Impedance_BCG.imag)
                print('Apparent_Impedance_BCG: %f+i%f' % (Apparent_Impedance_BCG.real,Apparent_Impedance_BCG.imag))
                print('Apparent_Reactance_BCG: %f' % Apparent_Reactance_BCG)
                Section_Detection_Function(Apparent_Reactance_BCG)
            elif f == 6 :
                Numerator_CAG = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0))))
                Denominator_CAG = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0))))
                Apparent_Impedance_CAG = Numerator_CAG / Denominator_CAG
                Apparent_Reactance_CAG = abs(Apparent_Impedance_CAG.imag)
                print('Apparent_Impedance_CAG: %f+i%f' % (Apparent_Impedance_CAG.real,Apparent_Impedance_CAG.imag))
                print('Apparent_Reactance_CAG: %f' % Apparent_Reactance_CAG)
                Section_Detection_Function(Apparent_Reactance_CAG)
            elif f == 7 :
                Numerator_AB = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1))))
                Denominator_AB = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1))))
                Apparent_Impedance_AB = Numerator_AB / Denominator_AB
                Apparent_Reactance_AB = abs(Apparent_Impedance_AB.imag)
                print('Apparent_Impedance_AB: %f+i%f' % (Apparent_Impedance_AB.real,Apparent_Impedance_AB.imag))
                print('Apparent_Reactance_AB: %f' % Apparent_Reactance_AB)
                Section_Detection_Function(Apparent_Reactance_AB)
            elif f == 8 :
                Numerator_BC = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,1)),Voltage_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2))))
                Denominator_BC = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,1)),Current_SpreadSheet_Imag_Part.cell_value(i,1)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2))))
                Apparent_Impedance_BC = Numerator_BC / Denominator_BC
                Apparent_Reactance_BC = abs(Apparent_Impedance_BC.imag)
                print('Apparent_Impedance_BC: %f+i%f' % (Apparent_Impedance_BC.real,Apparent_Reactance_BC.imag))
                print('Apparent_Reactance_BC: %f' % Apparent_Reactance_BC)
            elif f == 9 :
                Numerator_CA = ((complex((Voltage_SpreadSheet_Real_Part.cell_value(i,2)),Voltage_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Voltage_SpreadSheet_Real_Part.cell_value(i,0)),Voltage_SpreadSheet_Imag_Part.cell_value(i,0))))
                Denominator_CA = ((complex((Current_SpreadSheet_Real_Part.cell_value(i,2)),Current_SpreadSheet_Imag_Part.cell_value(i,2)))-(complex((Current_SpreadSheet_Real_Part.cell_value(i,0)),Current_SpreadSheet_Imag_Part.cell_value(i,0))))
                Apparent_Impedance_CA = Numerator_CA / Denominator_CA
                Apparent_Reactance_CA = abs(Apparent_Impedance_CA.imag)
                print('Apparent_Impedance_CA: %f+i%f ' % (Apparent_Impedance_CA.real,Apparent_Impedance_CA.imag))
                print('Apparent_Reactance_CA: %f' % Apparent_Reactance_CA)
                Section_Detection_Function(Apparent_Reactance_CA)
            elif f == 10 :
                Numerator_ABCG = cmath.rect((Voltage_Sequence_Magnitude.cell_value(i,0)),(Voltage_Sequence_Angle.cell_value(i,0)))
                Denominator_ABCG = cmath.rect((Current_Sequence_Magnitude.cell_value(i,0)),(Current_Sequence_Angle.cell_value(i,0)))
                Apparent_Impedance_ABCG = Numerator_ABCG / Denominator_ABCG
                Apparent_Reactance_ABCG = abs(Apparent_Impedance_ABCG.imag)
                print('Apparent_Impedance_ABCG: %f+i%f' % (Apparent_Impedance_ABCG.real,Apparent_Impedance_ABCG.imag))
                print('Apparent_Reactance_ABCG: %f' % Apparent_Reactance_ABCG)
                Section_Detection_Function(Apparent_Reactance_ABCG)
else:
    print('Number of Voltage and Current values not equal')

print('test to see if git is working')