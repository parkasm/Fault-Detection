# The fault type is determined from the fault current phasors using the logic based on Ratan Das Paper (Page 43)


"""Dictionary for the fault type and their respective return values
faultdict =	{
  "No fault": "0",
  "A-G": "1",
  "B-G": "2",
  "C-G": "3",
  "A-B-G" : "4",
  "B-C-G" : "5",
  "C-A-G" : "6",
  "A-B" : "7",
  "B-C" : "8",
  "C-A" : "9",
  "A-B-C-G" : "10"
}"""

# Function defined to get the fault type. The Input to the function are the magnitude of the current for all three phases at the measuring point
def fault_type(I_a, I_b, I_c):

    ##I_a = float(input("Enter Phase A current at measuring point: "))
    ##I_b = float(input("Enter Phase B current at measuring point: "))
    ##I_c = float(input("Enter Phase C current at measuring point: "))
    
    # Threshold current. The Value for this needs to come from either the utility company or have to be calculated somehow.  
    I_threshold = 150

    # Zero sequence component of the line currents.
    I_zero = (I_a + I_b + I_c)/3

    ## These conditional statements contains the logic to determine the type. Quantities involved are the current phasors, threshold current and the zero sequence current
    if abs(I_a) > abs(I_threshold):
        if abs(I_b) > abs(I_threshold):
            if abs(I_c) > abs(I_threshold):
                print('its a balanced 3 phase fault')
                return 10
            else:
                if abs(I_zero) > abs(I_threshold):
                    print('its a A-B-G fault')
                    return 4
                else:
                    print('its a A-B fault')
                    return 7
        else:
            if abs(I_c) > abs(I_threshold):
                if abs(I_zero) > abs(I_threshold):
                    print('its a C-A-G fault')
                    return 6
                else:
                    print('its a C-A fault')
                    return 9
            else:
                print('its a A-G fault')
                return 1
    else:
        if abs(I_b) > abs(I_threshold):
            if abs(I_c) > abs(I_threshold):
                if abs(I_zero) > abs(I_threshold):
                    print('its a B-C-G fault')
                    return 5
                else:
                    print('its a B-C fault')
                    return 8
            else:
                print('its a B-G fault')
                return 2
        else:
            if abs(I_c) > abs(I_threshold):
                print('its a C-G fault')
                return 3
            else:
                print('System has no fault')
                return 0

