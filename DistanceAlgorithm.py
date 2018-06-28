##THIS IS BASED ON THE DOCUMENT NAMED PPT-FAULT LOCATION ALGORITHM. USING THE SIMPLE IMPEDANCE EQUATIONS ON PAGE 4
#### REALLY IMPORTANT INFORMATION : THIS PAPER ASSUMES THAT THE FAULT RESISTANCE IS 0
from fault_type import fault_type
import xlrd
import panda
import numpy as np
import cmath
import math
import sys

f = fault_type(136.9,165.9,244.6)

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
Total_Number_of_Line_Sections = int(input('Enter the Total Number of Sections Line is divided into:')) ## THIS NUMBER IS BASED ON Simulink_Model.slx file
Total_Length_of_Line = Section_Length * Total_Number_of_Line_Sections

Zero_Sequence_Total_Line_Resistance = Zero_Seq_Resistace_ohm_per_km * Total_Length_of_Line
Zero_Sequence_Total_Line_Reactance = Zero_Seq_Inductance_Henry_per_km * 2 * math.pi * 60 * Total_Length_of_Line

Pos_Sequence_Total_Line_Resistance = Pos_Seq_Resistace_ohm_per_km * Total_Length_of_Line
Pos_Sequence_Total_Line_Reactance = Pos_Seq_Inductance_Henry_per_km * 2 * math.pi * 60 * Total_Length_of_Line

Zero_Sequence_Total_Line_Impedance = complex(Zero_Sequence_Total_Line_Resistance,Zero_Sequence_Total_Line_Reactance)
Pos_Sequence_Total_Line_Impedance = complex(Pos_Sequence_Total_Line_Resistance, Pos_Sequence_Total_Line_Reactance)

#print(Zero_Sequence_Total_Line_Impedance, Pos_Sequence_Total_Line_Impedance)

Ground_Compensation_Factor = ((Zero_Sequence_Total_Line_Impedance - Pos_Sequence_Total_Line_Impedance) / (3 * Pos_Sequence_Total_Line_Impedance))
print(Ground_Compensation_Factor)

##Loading the Excel Workbooks containing all the current, voltage and sequence component information
workbook = xlrd.open_workbook('Simulation_Data.xls')
workbook_2 = xlrd.open_workbook('Sequence_Components.xls')

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


#if(Voltage_SpreadSheet.nrows == Current_SpreadSheet.nrows):
    #Loop through all the iterations of the voltage/current phasors to get the apparent reactance for each of them
    #if f == 1 or f == 2 or f == 3:
Residual_Current_Real_Part = 0
Residual_Current_Imag_Part = 0
for j in range(0,Voltage_SpreadSheet.nrows):
    for k in range(0,Current_SpreadSheet.ncols):
        Residual_Current_Real_Part = Residual_Current_Real_Part + (Current_SpreadSheet_Real_Part.cell_value(j,k))
        Residual_Current_Imag_Part = Residual_Current_Imag_Part + (Current_SpreadSheet_Imag_Part.cell_value(j,k))
    Residual_Current = complex(Residual_Current_Real_Part,Residual_Current_Imag_Part)
    if f == 1:
        temp_variable_1_AG = (complex((Voltage_SpreadSheet_Real_Part.cell_value(j,0)),Voltage_SpreadSheet_Imag_Part.cell_value(j,0))) * Total_Length_of_Line
        temp_variable_2_AG = complex((Current_SpreadSheet_Real_Part.cell_value(j,0)),Current_SpreadSheet_Imag_Part.cell_value(j,0))
        temp_variable_3_AG = Residual_Current * Ground_Compensation_Factor
        Distance_AG = ((temp_variable_1_AG) / (Pos_Sequence_Total_Line_Impedance *(temp_variable_2_AG + temp_variable_3_AG)))
        print Distance_AG
    elif f == 2:
        temp_variable_1_BG = (complex((Voltage_SpreadSheet_Real_Part.cell_value(j,1)),Voltage_SpreadSheet_Imag_Part.cell_value(j,1))) * Total_Length_of_Line
        temp_variable_2_BG = complex((Current_SpreadSheet_Real_Part.cell_value(j,1)),Current_SpreadSheet_Imag_Part.cell_value(j,1))
        temp_variable_3_BG = Residual_Current * Ground_Compensation_Factor
        Distance_BG = ((temp_variable_1_BG) / (Pos_Sequence_Total_Line_Impedance *(temp_variable_2_BG + temp_variable_3_BG)))
        print Distance_BG
    elif f == 3:
        temp_variable_1_CG = (complex((Voltage_SpreadSheet_Real_Part.cell_value(j,2)),Voltage_SpreadSheet_Imag_Part.cell_value(j,2))) * Total_Length_of_Line
        temp_variable_2_CG = complex((Current_SpreadSheet_Real_Part.cell_value(j,2)),Current_SpreadSheet_Imag_Part.cell_value(j,2))
        temp_variable_3_CG = Residual_Current * Ground_Compensation_Factor
        Distance_CG = ((temp_variable_1_CG) / (Pos_Sequence_Total_Line_Impedance *(temp_variable_2_CG + temp_variable_3_CG)))
        print Distance_CG

    Residual_Current_Real_Part = 0
    Residual_Current_Imag_Part = 0
