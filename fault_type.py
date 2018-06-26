# The fault type is determined from the fault current phasors using the logic based on R Das Paper (Page 43)
# These are temporary value. In reality these number will be provided by user or fed from a file


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

def fault_type(I_a, I_b, I_c):

    ##I_a = float(input("Enter Phase A current at measuring point: "))
    ##I_b = float(input("Enter Phase B current at measuring point: "))
    ##I_c = float(input("Enter Phase C current at measuring point: "))
    I_threshold = 150
    I_zero = (I_a + I_b + I_c)/3

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

