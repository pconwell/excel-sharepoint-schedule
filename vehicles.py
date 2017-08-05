################################################################################
# Vehicle Assignments
################################################################################
# This is still a pretty terrible solution, but I haven't been able to figure
# out a better system, so here we are...

def assign(emp_num, eno, zone, w_date, config):

    # Supervisors
    S1 = config['Supervisors']['S1']
    S2 = config['Supervisors']['S2']

    # Police Vehicles
    V1 = config['PO_Vehicles']['V1']
    V2 = config['PO_Vehicles']['V2']
    V3 = config['PO_Vehicles']['V3']
    V4 = config['PO_Vehicles']['V4']
    V5 = config['PO_Vehicles']['V5']
    V6 = config['PO_Vehicles']['V6']
    V7 = config['PO_Vehicles']['V7']
    V8 = config['PO_Vehicles']['V8']

    # Bicycles
    B1 = config['Bicycles']['B1']

    # CSO Vehicles

    C1 = config['CSO_Vehicles']['C1']
    C2 = config['CSO_Vehicles']['C2']
    C3 = config['CSO_Vehicles']['C3']
    C4 = config['CSO_Vehicles']['C4']
    C5 = config['CSO_Vehicles']['C5']
    C6 = config['CSO_Vehicles']['C6']
    C7 = config['CSO_Vehicles']['C7']
    C8 = config['CSO_Vehicles']['C8']

    XX = C7

    # This 'array' assigns a vehicle (above) to an officer (emp_num), which is
    # according to the order in which an employee appears in the po_num and
    # cso_num lists. i.e. if '507' is the fisrt ENO in the po_num list, then
    # then emp_num[0] is '507'.

    patrol = {emp_num[0]: ['', '', '', B1, B1, B1, B1], # Greeks
              emp_num[1]: ['', '', '', B1, B1, B1, B1],

              emp_num[2]: ['', '', '', S1, S1, S1, S1], # Lieutenant

              emp_num[3]: ['', '', S1, S2, S2, S2, ''], # Sergeants
              emp_num[4]: [S1, S1, S2, '', '', '', S2],

              emp_num[5]: ['', '', '', V4, V1, V1, V1], # Officers
              emp_num[6]: [V2, '', '', '', V4, V2, V2],
              emp_num[7]: [V3, V3, '', '', '', V4, V3],
              emp_num[8]: [V5, V5, '', '', '', V5, V5],
              emp_num[9]: [V4, V4, V4, '', '', '', V4],
             emp_num[10]: [V1, V1, V1, V1, '', '', ''],
             emp_num[11]: ['', V2, V2, V2, V2, '', ''],
             emp_num[12]: ['', '', V3, V3, V3, V3, ''],

             emp_num[13]: ['', '', '', C1, C1, C1, C1], # CSOs
             emp_num[14]: [C2, '', '', '', C2, C2, C2],
             emp_num[15]: [C3, C3, '', '', '', C3, C3],
             emp_num[16]: [C4, C4, C4, '', '', '', C4],
             emp_num[17]: [C5, C5, C5, '', '', '', C5], #Stewart
             emp_num[18]: [C1, C1, C1, XX, '', '', ''], #Dixon
             emp_num[19]: ['', C2, C2, C2, XX, '', ''],
             emp_num[20]: ['', '', C3, C3, C3, XX, ''],
             emp_num[21]: ['', '', XX, C4, C4, C4, ''],
             emp_num[22]: ['', '', '', C5, C5, C5, XX]
             }

    v = patrol[eno_lookup(emp_num, eno)][weekday_num(w_date)]

    if zone == 20:
        v = S1
    elif zone == 66:
        v = C6

# The try:except block catches entries on the schedule that are not 'zones'
# e.g. "X" is not a number so it will throw an error.
    try:
        if v == '' and int(str(zone)[:2]) < 30:
            v = V6
        elif v == '' and int(str(zone)[:2]) > 30:
            v = C7
    except ValueError:
        v = ''

    return v

# Fix for isoweekday, which returns Sunday as '7'
# This will now return Sunday = 0, Saturday = 6
def weekday_num(day_num):
    if day_num.isoweekday() == 7:
        return 0
    else:
        return day_num.isoweekday()

# Finds the order for employee numbers in the emp_num list
# In other words if '507' is the first employee number on the list
# '507' is emp_num[0]
def eno_lookup(emp_num, eno):
    return emp_num[emp_num.index(str(eno))]
