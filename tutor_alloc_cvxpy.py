"""
Tutor allocation program for university courses
Created: 20/07/2022
Modified: 14/03/2023
Authors: Liam Timms - liam.timms@uq.edu.au

INSTRUCTIONS
--- Installation ---
You may need to install the packages cvxpy. Instructions for doing are on https://www.cvxpy.org/install/index.html

--- The Excel spreadsheet ---
Workshop times & column names:
  -   1st column should be called 'Full name'
  -   All other columns should be the workshop times. Workshop times are labelled "Day StartTime-EndTime Suffix"
      Internal workshops for the main course don't need a suffix. External workshops should end with 'EX'.
      Workshops for additional courses (e.g. SCIE1100) should have the course code as a suffix.

  -   If there are multiple workshop on in a given timeslot, duplicate the column and add the workshop room to the
      column name, e.g. Monday 2pm-4pm ILC1 and Monday 2pm-4pm ILC3.

Tutor names
  -   Add (Super) to the end of supertutors' names

Spreadsheet entries
  -   Tutor availabilities: Entries should be 'Available', 'IfNeeded', or 'NotAvailable' (no spaces)

  -   Add a row for how many tutors are assigned to each workshop. Name should end in 'tutors', e.g. 'Num tutors'.
      Avoid having any other rows contain the substring 'tutors'.

Spreadsheet sheet
  -   The first sheet of the Excel spreadsheet is named 'Availability' and contains
      workshop availabilities and tutor numbers.

  -   The second sheet is named 'Allocations' and contains how many workshops are assigned to each tutor
      (one column for each course), as well as tutors' experience. The experience column is labelled 'Experience'
      and has entries 1 for experienced, and 0 otherwise.

  -   The third sheet is named 'Conflicts'. There are two columns of entries, labelled 'Tutor 1' and 'Tutor 2'.
      Each row after that contains pairs of tutors that cannot be in the same workshop.

Debugging
  -   If you run the program and the result it "Unable to retrieve attribute 'x'", then the timetable is infeasible.
      Try commenting out the last constraint and run the program again. If it's still infeasible, uncomment that
      constraint and comment out the second last constraint. Repeat until the program is feasible - the constraint you
      commented out that time is likely the one causing the timetable to be infeasible. Think about what's in your
      data that could cause problems with that constraint. For example, maybe the tutors' availabilities are too
      restrictive, or there isn't enough flexibility to schedule a supertutor on the first day of workshops.
      Note: the constraints appear after the line "# --------- The constraints ---------", and look like
      variable = {i: m.addConstr( ... ) ...} etc. To get rid of one constraint, comment out everything from the
      "variable = {" down to the closing brace "}"
"""
from random import seed
import cvxpy as cp
from pandas import DataFrame, read_excel, read_csv
import numpy as np
from itertools import combinations


def import_spreadsheet(fname, sname, blank_value):
    """
    :param fname: Name of the Excel spreadsheet file, including path
    :param sname: Name of the sheet in the Excel file
    :param blank_value: value/string to replace any blank values in the spreadsheet with
    :return: (pandas.dataframe) dataframe of Excel spreadsheet contents
    """

    # Read in the file as a dataframe. index_col=0 uses the 0th column (tutor names) as the row names
    # Use read_excel() if the file extension is one of the supported extensions
    if any([file_name.endswith(x) for x in ['xls', 'xlsx', 'xlsm', 'xlsb', 'odf', 'ods', 'odt']]):
        df = read_excel(fname, sheet_name=sname, index_col=0)

    # Use read_csv
    elif file_name.endswith('.csv'):
        df = read_csv(fname, sheet_name=sname, index_col=0)

    else:
        raise ValueError('The specified file is not a supported file type. Supported types are xls, xlsx, xlsm,'
                         ' xlsb, odf, ods, odt, and csv.')
    # Replace blank values with blank_value
    df.fillna(value=blank_value, inplace=True)
    # Remove any whitespace at the start and end of the column headers
    df.rename(mapper=str.strip, axis='columns', inplace=True)

    return df


def yes_no_question(question_text):
    """
    :param question_text: (str) text to display with the input() function
    :return: (str) 'yes' or 'no', the user's answer
    """
    # Ask the user for an input
    answer = input(question_text)

    # If the user does not input a valid answer
    while answer.lower().strip() not in ['yes', 'no']:
        answer = input("Incorrect input. Enter either 'yes' or 'no':")

    return answer


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

file_name = input('Enter the file name of the tutor workshop availability Excel spreadsheet'
                  ' (including path, use \\\ instead of single \)'
                  ' \nFor example, "Sem 1 2023 resources\\\SCIE1000_1100_availabilities.xlsx"')

# --------- READING IN THE EXCEL SPREADSHEETS ---------
# Read in the availabilities spreadsheet as a dataframe, specify 1st sheet
workshop_avail_df = import_spreadsheet(file_name, sname='Availability', blank_value='NotAvailable')

# Dataframe for how many workshops assigned to each tutor & tutors' experience
workshop_num_df = import_spreadsheet(file_name, sname='Allocations', blank_value=0)

# Split the above dataframe (df) into a df containing tutors' experience and gender identity, and a df with workshop
# number allocations
workshop_exp_df = workshop_num_df[['Experience', 'Gender ID']]

# Remove the 'Experience' column from the workshop numbers df
# "axis=1" refers to columns. Axis 0 would be rows.
workshop_num_df = workshop_num_df.drop(labels=['Experience', 'Gender ID'], axis=1)

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

# Dictionary of diversity indicator: equals 1 if tutor i and j have different gender identities, 0 otherwise
Div_ij = {(i, j): (1 if workshop_exp_df.loc[i]['Gender ID'] is not workshop_exp_df.loc[j]['Gender ID'] else 0)
          for i in Tutors for j in Tutors if i is not j}

# Number of tutors assigned to each workshop
num_tutors_row = find_row_name('tutors', workshop_avail_df)  # First find the row name containing the tutor numbers
N_w = workshop_avail_df.loc[num_tutors_row]

# Number of workshops assigned to each tutor i: [SCIE1000, SCIE1100]
# e.g. M_i[tutor][0] = SCIE1000,  M_i[tutor][1] = SCIE1100
M_i = {tutor: [alloc for alloc in workshop_num_df.loc[tutor]] for tutor in Tutors}

# List of tutor conflicts. Tuples in this list are conflicting pairs of tutors.
# Tutor relationship pairs fall in this category.
# e.g. C_ij[0] = [tutor X_iw, tutor Y]
conflicts = yes_no_question("Are there any tutor conflicts (yes/no):")

if conflicts.lower() == 'yes':
    # Read in the tutor conflicts
    workshop_conflict_df = read_excel('Sem 1 2023 resources\\SCIE1000_1100_availabilities.xlsx',
                                      sheet_name='Conflicts')

    # Create list of tutor conflicts. Each element will be a list [Tutor 0, Tutor 1]
    C_ij = [  # List of the tutors in conflict given in row k of Excel sheet
        [tutor for tutor in workshop_conflict_df.loc[k]]
        for k in workshop_conflict_df.index]

# Are there any SCIE1100 workshops to schedule?
do_scie1100 = yes_no_question("Are you scheduling SCIE1100 as well as SCIE1000? (yes/no)")

# Determine weighting for gender diversity in the objective function
weight = float(input("What is the weighting (w) for gender diverse tutoring allocations?\n"
                     "0 < w < 1 means that tutors' workshop preferences are weighted more than gender diversity."
                     " Conversely for w > 1. \n"
                     "Enter value of w: "))
if weight < 0:
    weight = 0
    print("Weighting entered was negative. Weight has been set to 0.")

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

# --------- The variables ---------
# X_iw=1 if tutor i is allocated to workshop w, 0 otherwise. Only create variable if tutor can take workshop.
X_iw = {(i, w): cp.Variable(boolean=True) for i in Tutors for w in Workshops if P_iw[i][w] != 0}

# Y_ijw = 1 if tutors i and j are both allocated to workshop w, 0 otherwise (for workshops that require 2 tutors)
Y_ijw = {(i, j, w): cp.Variable(boolean=True) for (i, j) in combinations(Tutors, 2) for w in Workshops
         if N_w[w] == 2 if (i, w) in X_iw if (j, w) in X_iw}

# Z_ijkw = 1 if tutors i, j, and k are all allocated to workshop w, 0 otherwise (for workshops that require 3 tutors)
Z_ijkw = {(i, j, k, w): cp.Variable(boolean=True) for (i, j, k) in combinations(Tutors, 3) for w in Workshops
          if N_w[w] == 3 if (i, w) in X_iw if (j, w) in X_iw if (k, w) in X_iw}

# --------- The objective ---------
# Maximise tutor preferences, trying to avoid 'if needed' allocations, with a bonus for having high average gender
# diversity in workshop allocations.
# Max value for sum of preferences is sum(N_w), so dividing by sum(N_w) normalises the preference term in the objective.
# Max value for sum of gender diversity in 2-tutor workshops is the no. of workshops that require 2 tutors.
# For 3-tutor workshops, the max value is 3 x the no. of workshops requiring 3 tutors. Dividing the gender diversity
# term by sum(N_w==2) + 3 x sum(N_w==3) then normalises the average gender diversity.
objective = cp.Maximize(
    1 / Available / np.sum(N_w) * sum(P_iw[i][w] * X_iw[i, w] for i in Tutors for w in Workshops if (i, w) in X_iw) \
    + weight / (np.sum(N_w == 2) + 3 * np.sum(N_w == 3)) * (sum(Y_ijw[i, j, w] * Div_ij[i, j] for (i, j, w) in Y_ijw)
    + sum(Y_ijw[i, j, w] * Div_ij[i, j] for (i, j, w) in Y_ijw))
)

# --------- The constraints ---------
# Initialise list to contain all the constraints
constraints = []
# Each workshop must have correct number of tutors teaching it
constraints += [sum(X_iw[i, w] for i in Tutors if (i, w) in X_iw) == N_w[w] for w in Workshops]

if do_scie1100.lower() == 'no':
    # Make sure each tutor is allocated the correct number of workshops for SCIE1000
    # M_i[i] will only have one element, since only SCIE1000 is being run this semester.
    constraints += [sum(X_iw[i, w] for w in Workshops if (i, w) in X_iw) == M_i[i][0] for i in Tutors]

elif do_scie1100.lower() == 'yes':
    # Make sure each tutor is allocated the correct number of workshops for SCIE1000
    # M_i[i] has two elements: M_i[i][0] -> SCIE1000, M_i[i][1] -> SCIE1100
    constraints += [sum(X_iw[i, w] for w in Workshops if (i, w) in X_iw if '1100' not in Time_slots[w]) == M_i[i][0]
                    for i in Tutors]

    # Make sure each tutor is allocated the correct number of workshops for SCIE1100
    constraints += [sum(X_iw[i, w] for w in Workshops if (i, w) in X_iw if '1100' in Time_slots[w]) == M_i[i][1]
                    for i in Tutors]

else:
    # If do_scie1100 != 'yes' and !='no', then something has gone wrong :(
    raise ValueError('Invalid input for whether or not SCIE1100 is running this semester.')

# Tutors can teach at most one workshop at a time -> sum over all workshops on same day that overlap with workshop w
# If the workshops start within an hour of each other, they will overlap, provided they are on the same day
constraints += [sum(X_iw[i, v] for v in Workshops if (i, v) in X_iw
                    if abs(Workshop_time[v][0] - Workshop_time[w][0]) <= 100
                    if Workshop_day[v] == Workshop_day[w]) <= 1
                for i in Tutors for w in Workshops]

# At least one experienced tutor per workshop (assuming that there are at least two tutors per workshop)
# A tutor is experienced if their experience in the Excel sheet 'Allocations' = 1
# If there are workshops with only one tutor, this constraint can be edited to: "workshop_exp_df.loc[i] == 0) <= 1"
# to allow for inexperienced tutors tutoring by themselves (unlikely)
constraints += [sum(X_iw[i, w] for i in Tutors if (i, w) in X_iw if workshop_exp_df.loc[i]['Experience'] == 1)
                >= 1 for w in Workshops
                ]

# If there are any conflicts
if conflicts.lower() == 'yes':
    # Tutors with conflicts cannot teach together. Note: C_ij contains lists ij = [Tutor i, Tutor j]
    constraints += [X_iw[i, w] + X_iw[j, w] <= 1
                    for i, j in C_ij for w in Workshops if (i, w) in X_iw and (j, w) in X_iw]

# A supertutor is ideally teaching a workshop on the first day of workshops during the week.
# This constraint can be removed if it makes the timetable infeasible.
constraints += [sum(X_iw[i, w] for w in Workshops if (i, w) in X_iw if first_workshop in Time_slots[w]) >= 1
                for i in Supertutors]

# Supertutors shouldn't teach the same workshop - inefficient use of resources
# This constraint can be removed if it makes the timetable infeasible.
constraints += [sum(X_iw[i, w] for i in Supertutors if (i, w) in X_iw) <= 1
                for w in Workshops]

# Constraints for Y_ijw and Z_ijkw
# We want Y_ijw = 1 if and only if X_iw = X_jw = 1 (see comment at definition of Y_ijw)
for (i, j, w) in Y_ijw:
    constraints += [Y_ijw[i, j, w] <= X_iw[i, w],  # Y_ijw must be 0 if X_iw is 0 ((remember that Y_ijw is binary)
                    Y_ijw[i, j, w] <= X_iw[j, w],  # # Y_ijw must be 0 if X_jw is 0
                    Y_ijw[i, j, w] >= X_iw[i, w] + X_iw[j, w] - 1  # Y_ijw must be 1 if both X_iw and X_jw are 1
                    ]

for (i, j, k, w) in Z_ijkw:
    # Z_ijkw must be 0 if any of X_iw, X_jw, X_kw are 0
    constraints += [Z_ijkw[i, j, k, w] <= X_iw[i, w],
                    Z_ijkw[i, j, k, w] <= X_iw[j, w],
                    Z_ijkw[i, j, k, w] <= X_iw[k, w],
                    # Z_ijkw must be 1 if X_iw, X_jw, and X_kw are all 1
                    Z_ijkw[i, j, k, w] >= X_iw[i, w] + X_iw[j, w] + X_iw[k, w] - 2
                    ]

# ------------- Manual constraints - tutor preferences -------------
# Tutors can be manually scheduled by using the following line of code (replace occurrences of TUTOR with tutor's name,
# and replace TIMESLOT with the name of the workshop as is appears in the list Time_slots)

# TUTORPreference = m.addConstr(X_iw['TUTOR', Time_slots.index('TIMESLOT')] == 1)

# Alternatively, you can include tutors' general preferences for a specific day or time with the following line of code
# (replace occurrences of TUTOR with tutor's name, and DETAIL with the specific day or time, e.g. 'Mon' or '8am')
# TUTORPreference = m.addConstr(sum(X_iw['TUTOR', w] for w in Workshops if DETAIL in Time_slots[w]) >= 1)

problem = cp.Problem(objective, constraints)
problem.solve()
#
# # Save the results as a dataframe. For each workshop, mark the allocated tutors with an X_iw.
results_df = DataFrame(index=np.array(Tutors), columns=np.array(Time_slots))
for (i, w) in X_iw:
    if X_iw[i, w].value > 0.9:
        print()
        results_df.loc[i, Time_slots[w]] = 'X'

# Export the allocation dataframe to an Excel file
# results_df.to_excel('tutor_workshop_schedule_cvxpy.xlsx', sheet_name='Timetable', index=True)
