"""
SCIE1000/1100 tutor allocation program
Created: 20/07/2022
Modified: 07/02/2023
Authors: Liam Timms
"""
from random import seed
from gurobipy import Model, GRB, quicksum
from pandas import DataFrame, read_excel
import numpy as np
import re


def find_row_name(substring, df):
    """
    Find the first row in a given dataframe that contains the given substring.
    find_row_name('tut', my_df) >>> 'Tutor names'

    :param substring: (str) string that is contained in the target row
    :param df: (pandas dataframe) the dataframe containing the rows you're searching through
    :return: (str) name of first row containing the substring
    """

    row_matches = [row for row in df.index if substring in row]

    return row_matches[0]


# Set the random seed for the model solver (Gurobi)
seed(42)

# INSTRUCTIONS FOR CREATING THE EXCEL SPREADSHEET OF AVAILABILITIES
#
# Workshop times & column names:
#   -   1st column should be called 'Full name'
#   -   All other columns should be the workshop times. Workshop times are labelled "Day StartTime-EndTime Suffix"
#       Internal workshops for the main course don't need a suffix. External workshops should end with 'EX'.
#       Workshops for additional courses (e.g. SCIE1100) should have the course code as a suffix.
#
#   -   If there are multiple workshop on in a given timeslot, duplicate the column and add the workshop room to the
#       column name, e.g. Monday 2pm-4pm ILC1 and Monday 2pm-4pm ILC3.
#
#   -   Make sure there's no extra whitespace in the column names, e.g. a space at the start/end of 'Monday 2pm-4pm'
#
# Tutor names
#   -   Add (Super) to the end of supertutors' names
#
# Spreadsheet entries
#   -   Tutor availabilities: Entries should be 'Available', 'IfNeeded', or 'NotAvailable' (no spaces)
#
#   -   Add a row for how many tutors are assigned to each workshop. Name should end in 'tutors', e.g. 'Num tutors'.
#       Avoid having any other rows contain the substring 'tutors'.
#
# Spreadsheet sheet
#   -   The first sheet of the Excel spreadsheet is named 'Availability' and contains
#       workshop availabilities and tutor numbers.
#       The second sheet is named 'Allocations' and contains how many workshops are assigned to each tutor
#       (one column for each course), as well as tutors' experience. The experience column is labelled 'Experience'
#       and has entries 1 for experienced, and 0 otherwise.
#
#   -   The third sheet is named 'Conflicts'. There are two columns of entries, labelled 'Tutor 1' and 'Tutor 2'.
#       Each row after that contains pairs of tutors that cannot be in the same workshop.
#
# Debugging
#   -   If you run the program and the result it "Unable to retrieve attribute 'x'", then the timetable is infeasible.
#       Try commenting out the last constraint and run the program again. If it's still infeasible, uncomment that
#       constraint and comment out the second last constraint. Repeat until the program is feasible - the constraint you
#       commented out that time is likely the one causing the timetable to be infeasible. Think about what's in your
#       data that could cause problems with that constraint. For example, maybe the tutors' availabilities are too
#       restrictive, or there isn't enough flexibility to schedule a supertutor on the first day of workshops.
#       Note: the constraints appear after the line "# --------- The constraints ---------", and look like
#       variable = {i: m.addConstr( ... ) ...} etc. To get rid of one constraint, comment out everything from the
#       "variable = {" down to the closing brace "}"


file_name = input('Enter the file name of the tutor workshop availability Excel spreadsheet'
                  ' (including path, use \\\ instead of single \)'
                  ' \nFor example, "Sem 1 2023 resources\\\SCIE1000_1100_availabilities.xlsx"')

# --------- READING IN THE EXCEL SPREADSHEETS ---------
# Read in the availabilities spreadsheet as a dataframe, specify 1st sheet
# index_col=0 uses the 0th column (tutor names) as the row names
workshop_avail_df = read_excel(file_name, sheet_name='Availability',
                               index_col=0)
# Replace any blank values (read in as NaN) with 'NotAvailable'. "inplace=True" modifies the existing dataframe.
workshop_avail_df.fillna(value='NotAvailable', inplace=True)

# Dataframe for how many workshops assigned to each tutor & tutors' experience
workshop_num_df = read_excel(file_name, sheet_name='Allocations',
                             index_col=0)
# Replace blank values with 0 (no workshops assigned)
workshop_num_df.fillna(value=0, inplace=True)

# Split the above dataframe (df) into a df containing tutors' experience, and a df with workshop number allocations
workshop_exp_df = workshop_num_df['Experience']
# Remove the 'Experience' column from the workshop numbers df
workshop_num_df = workshop_num_df.drop(labels='Experience', axis=1)  # "axis=1" refers to columns. Axis 0 would be rows.

# --------- THE SETS ---------
# List of tutors (ignore any rows whose name ends in 'tutors', since that row should be the tutor allocation numbers
Tutors = [t for t in workshop_avail_df.index if not t.lower().endswith('tutors')]

# List of supertutors (at least one should be teaching on the 1st day of workshops each week)
Supertutors = [tutor for tutor in Tutors if tutor.lower().endswith('(super)')]

# List of times when workshops are scheduled (these are the dataframe columns).
# Assumed that workshops are 2 hours long.
Time_slots = [time for time in workshop_avail_df.columns]

# Set of workshops scheduled for this semester
# Listed in order given in availability spreadsheet
Workshops = range(len(Time_slots))

# Dictionary of the day that workshop w is on.
# Split the workshop name by spaces, then take the 0th result. This should be the day.
Workshop_day = {w: Time_slots[w].split(' ')[0] for w in Workshops}

# Initialise dictionary for the 24-hr time of each workshop
# Workshop_time[w][0] -> start time of workshop w. Workshop_time[w][1] -> end time of workshop w.
Workshop_time = {w: [] for w in Workshops}

# Workshop start and end times
for w in Workshops:
    workshop = Time_slots[w]

    # Split on the spaces, extract the element in position 1. Since all timeslots look like Day StartTime-EndTime Suffix
    time_period = workshop.split(' ')[1]
    # Then split into start and end times
    time_period = time_period.split('-')
    # Convert to 24-hr time
    time_period_24 = []

    for time in time_period:
        if 'am' in time.lower() or '12pm' in time.lower():
            # Remove the 'am' and convert to 24-hour time
            time_24 = int(time.lower().replace('am', '').replace('pm', '')) * 100

        elif 'pm' in time.lower():
            # Remove the 'pm' and convert to 24-hour time.
            time_24 = int(time.lower().replace('pm', '')) * 100 + 1200

        else:
            raise ValueError(f"Workshop {workshop} time doesn't contain 'am' or 'pm'.")

        time_period_24.append(time_24)

    # Add the time period of workshop w in 24-hr time to the dictionary
    Workshop_time[w] = time_period_24

# Create dictionary containing the list of workshops that overlap with workshop w (including workshop w)
# Assume that workshops only start on the hour, and all have a duration of 2 hours
Overlap_workshops = {w: [
    # List all workshops that start up to 1 hour before or up to 1 hour after workshop w, and are on the same day.
    v for v in Workshops if abs(Workshop_time[v][0] - Workshop_time[w][0]) <= 100 and Workshop_day[v] == Workshop_day[w]
] for w in Workshops}

# What day is the first workshop? Need this to schedule a supertutor on first day of workshops
first_workshop = None
for day in ['Mon', 'Tues', 'Wed', 'Thur', 'Fri']:
    # If the substring 'mon' appears in any workshop timeslot, then Monday is the first day of workshops.
    # If not Monday, then try Tuesday, etc.
    if any(day.lower() in workday.lower() for workday in Workshop_day.values()):
        # The first workshops falls on this day
        first_workshop = day
        break

# If we get through the loop without changing the value of first_workshop, then something has gone wrong
if first_workshop is None:
    raise ValueError('Either there are no workshops, or none of the workshop column names contain'
                     ' "Mon", "Tues", "Wed", "Thurs", or "Fri".')

# --------- THE DATA ---------
# Tutors' preferences for each workshop. "Available" weight is set here. Increase it to more strongly
# favour "Available" over "If Needed"
Available = 10
IfNeeded = 1
NotAvailable = 0

# The Availability dataframe entries are all strings. This alters the existing dataframe by
# replacing the strings with the variables Available, IfNeeded, and NotAvailable.
workshop_avail_df.replace(to_replace=['Available', 'IfNeeded', 'NotAvailable'],
                          value=[Available, IfNeeded, NotAvailable], inplace=True)

# Dictionary of tutors' availabilities. Each entry in the dictionary is a pandas.Series.
P_iw = {tutor: workshop_avail_df.loc[tutor] for tutor in Tutors}

# Number of tutors assigned to each workshop
num_tutors_row = find_row_name('tutors', workshop_avail_df)  # First find the row name containing the tutor numbers
N_w = workshop_avail_df.loc[num_tutors_row]

# Number of workshops assigned to each tutor i: [SCIE1000, SCIE1100]
# e.g. M_i[tutor][0] = SCIE1000,  M_i[tutor][1] = SCIE1100
M_i = {tutor: [alloc for alloc in workshop_num_df.loc[tutor]] for tutor in Tutors}

# List of tutor conflicts. Tuples in this list are conflicting pairs of tutors.
# Tutor relationship pairs fall in this category.
# e.g. C_ij[0] = [tutor X, tutor Y]
conflicts = input("Are there any tutor conflicts (yes/no):")
if conflicts.lower() == 'yes':
    # Read in the tutor conflicts
    workshop_conflict_df = read_excel('Sem 1 2023 resources\\SCIE1000_1100_availabilities.xlsx',
                                      sheet_name='Conflicts')

    # Create list of tutor conflicts. Each element will be a list [Tutor 0, Tutor 1]
    C_ij = [  # List of the tutors in conflict given in row k of Excel sheet
        [tutor for tutor in workshop_conflict_df.loc[k]]
        for k in workshop_conflict_df.index]

# --------- ERROR CHECKING ---------
# Make sure every tutor in the Workshop Availability Excel sheet is in the Workshop Allocation sheet
if len(Tutors) != len([t for t in workshop_avail_df.index if not t.lower().endswith('tutors')]):
    raise ValueError('The number of tutor entries in the Availability sheet is'
                     'not the same as the Workshop allocation sheet.')

# N_w is the no. of tutors assigned to workshop w. M_i is the no. of workshops assigned to tutor i.
# The total tutors needed to staff all the workshops should equal the total workshops assigned to all tutors.
if abs(np.sum(N_w) - np.sum([M_i[i] for i in M_i])) > 0.1:
    raise ValueError(f'The number of tutors needed to staff all workshops is not equal to the'
                     f' total number of workshops assigned to all tutors.')

# --------- THE MODEL ---------
m = Model('tutor_alloc')

# --------- The variables ---------
# X=1 if tutor i is allocated to workshop w, 0 otherwise. Only create variable if tutor can take workshop.
X = {(i, w): m.addVar(vtype=GRB.BINARY) for i in Tutors for w in Workshops if P_iw[i][w] != 0}

# --------- The objective ---------
# Maximise tutor preferences, trying to avoid 'if needed' allocations
m.setObjective(
    quicksum(P_iw[i][w] * X[i, w] for i in Tutors for w in Workshops if (i, w) in X)
    , GRB.MAXIMIZE)

# --------- The constraints ---------
# Each workshop must have correct number of tutors teaching it
WorkshopsStaffed = {w: m.addConstr(quicksum(X[i, w] for i in Tutors if (i, w) in X) == N_w[w])
                    for w in Workshops}

# Are there any SCIE1100 workshops to schedule?
do_scie1100 = input("Is SCIE1100 being run this semester? (yes/no)")
# If the user does not input
while do_scie1100.lower() not in ['yes', 'no']:
    do_scie1100 = input("Incorrect input. Enter either 'yes' or 'no':")

if do_scie1100.lower() == 'no':
    # Make sure each tutor is allocated the correct number of workshops for SCIE1000
    # M_i[i] will only have one element, since only SCIE1000 is being run this semester.
    NumWorkshops = {i: m.addConstr(quicksum(X[i, w] for w in Workshops if (i, w) in X) == M_i[i][0])
                    for i in Tutors}

elif do_scie1100.lower() == 'yes':
    # Make sure each tutor is allocated the correct number of workshops for SCIE1000
    # M_i[i] has two elements: M_i[i][0] -> SCIE1000, M_i[i][1] -> SCIE1100
    NumWorkshops = {i: m.addConstr(quicksum(X[i, w] for w in Workshops
                                            if (i, w) in X if '1100' not in Time_slots[w]) == M_i[i][0])
                    for i in Tutors}

    # Make sure each tutor is allocated the correct number of workshops for SCIE1100
    NumWorkshops1100 = {i: m.addConstr(quicksum(X[i, w] for w in Workshops
                                                if (i, w) in X if '1100' in Time_slots[w]) == M_i[i][1])
                        for i in Tutors}

else:
    # If do_scie1100 != 'yes' and !='no', then something has gone wrong :(
    raise ValueError('Invalid input for whether or not SCIE1100 is running this semester.')

# Tutors can teach at most one workshop at a time -> sum over all workshops on same day that overlap with workshop w
# If the workshops start within an hour of each other, they will overlap, provided they are on the same day
OnePlaceAtATime = {(i, w): m.addConstr(
    quicksum(X[i, v] for v in Workshops if (i, v) in X
             if abs(Workshop_time[v][0] - Workshop_time[w][0]) <= 100
             if Workshop_day[v] == Workshop_day[w]) <= 1
) for i in Tutors for w in Workshops}

# At least one experienced tutor per workshop (assuming that there are at least two tutors per workshop)
# A tutor is experienced if their experience in the Excel sheet 'Allocations' = 1
# If there are workshops with only one tutor, this constraint can be edited to: "workshop_exp_df.loc[i] == 0) <= 1"
# to allow for inexperienced tutors tutoring by themselves (unlikely)
AtMostOneInexp = {w: m.addConstr(
    quicksum(X[i, w] for i in Tutors if (i, w) in X if workshop_exp_df.loc[i] == 1) >= 1)
    for w in Workshops
}

# If there are any conflicts
if conflicts.lower() == 'yes':
    # Tutors with conflicts cannot teach together. Note: C_ij contains lists ij = [Tutor i, Tutor j]
    NoConflicts = {(i, j, w): m.addConstr(X[i, w] + X[j, w] <= 1)
                   for i, j in C_ij for w in Workshops if (i, w) in X and (j, w) in X}

# A supertutor is ideally teaching a workshop on the first day of workshops during the week.
# This constraint can be removed if it makes the timetable infeasible.
SupertutorWorkshop = {i: m.addConstr(
    quicksum(X[i, w] for w in Workshops if (i, w) in X if first_workshop in Time_slots[w]) >= 1
) for i in Supertutors}

# Supertutors shouldn't teach the same workshop - inefficient use of resources
# This constraint can be removed if it makes the timetable infeasible.
SupertutorOverlap = {w: m.addConstr(
    quicksum(X[i, w] for i in Supertutors if (i, w) in X) <= 1
) for w in Workshops}

# ------------- Manual constraints - tutor preferences -------------
# Tutors can be manually scheduled by using the following line of code (replace occurrences of TUTOR with tutor's name,
# and replace TIMESLOT with the name of the workshop as is appears in the list Time_slots)

# TUTORPreference = m.addConstr(X['TUTOR', Time_slots.index('TIMESLOT')] == 1)

# Alternatively, you can include tutors' general preferences for a specific day or time with the following line of code
# (replace occurrences of TUTOR with tutor's name, and DETAIL with the specific day or time, e.g. 'Mon' or '8am')
# TUTORPreference = m.addConstr(quicksum(X['TUTOR', w] for w in Workshops if DETAIL in Time_slots[w]) >= 1)

m.optimize()

# Save the results as a dataframe. For each workshop, mark the allocated tutors with an X.
results_df = DataFrame(index=np.array(Tutors), columns=np.array(Time_slots))
for (i, w) in X:
    if X[i, w].x > 0.9:
        results_df.loc[i, Time_slots[w]] = 'X'

# Export the allocation dataframe to an Excel file
results_df.to_excel('tutor_workshop_schedule.xlsx', sheet_name='Timetable', index=True)
