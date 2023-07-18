# Allocating tutors to course tutorials
Program and mathematical model for optimal allocation of university tutors to course tutorials.

tutor_alloc_gurobi.py and tutor_alloc_cvxpy.py implement a mixed-integer programming model for optimal tutor-workshop allocations, using Gurobi and CVXPY (open source) as solvers, respectively.
They read in data from a user-provided Excel spreadsheet, and are set up so the user is not required to interact with the underlying mathematical model.

Tutor_allocation_model.pdf describes the mixed-integer mathematical program that underpins the allocation code.

example_spreadsheet.xlsx provides a template/example of the spreadsheet containing tutor availabilities, workshop times, and tutor conflicts.

### Getting started
To allocate your tutors, download either tutor_alloc_cvxpy.py (open-source solver) or tutor_alloc_gurobi (if you have Gurobi). Create a spreadsheet of tutor availabilities, workshop times, and tutor conflicts in the same format as example_spreadsheet (detailed instructions below). Then you can run the program and it will allocate your tutors!

### Creating the spreadsheet of availabilities
Supported spreasheet file types are xls, xlsx, xlsm, xlsb, odf, ods, odt, and csv.


Workshop times & column names:
  -   1st column should be called 'Full name'
  -   All other columns should be the workshop times. Workshop times are labelled "Day StartTime-EndTime Suffix"
      Internal workshops for the main course don't need a suffix. External workshops should end with 'EX'.
      Workshops for additional courses (e.g. SCIE1100) should have the course code as a suffix.

  -   If there are multiple workshop on in a given timeslot, duplicate the column and add the workshop room to the
      column name, e.g. Monday 2pm-4pm ILC1 and Monday 2pm-4pm ILC3.

  -   Make sure there's no extra whitespace in the column names, e.g. a space at the start/end of 'Monday 2pm-4pm'

Tutor names
  -   Add (Super) to the end of supertutors' names

Spreadsheet entries
  -   Tutor availabilities: Entries should be 'Available', 'IfNeeded', or 'NotAvailable' (no spaces)

  -   Add a row for how many tutors are assigned to each workshop. Name should end in 'tutors', e.g. 'Num tutors'.
      Avoid having any other rows contain the substring 'tutors'.

Spreadsheet sheets
  -   The first sheet of the Excel spreadsheet is named 'Availability' and contains
      workshop availabilities and tutor numbers.
      
  -   The second sheet is named 'Allocations' and contains how many workshops are assigned to each tutor
      (one column for each course), as well as tutors' experience and gender identities.
      The experience column is labelled 'Experience' and has entries 1 for experienced, and 0 otherwise.
      The gender identities column is labelled 'Gender ID' and can have any entry as long as the entries are
      consistent, e.g. all tutors identifying as male are labelled 'Male'. 

  -   The third sheet is named 'Conflicts'. There are two columns of entries, labelled 'Tutor 1' and 'Tutor 2'.
      Each row after that contains pairs of tutors that cannot be in the same workshop.

Debugging (will fix this in future update)
  -   If you run the program and the result it "Unable to retrieve attribute 'x'", then the timetable is infeasible.
      Try commenting out the last constraint and run the program again. If it's still infeasible, uncomment that
      constraint and comment out the second last constraint. Repeat until the program is feasible - the constraint you
      commented out that time is likely the one causing the timetable to be infeasible. Think about what's in your
      data that could cause problems with that constraint. For example, maybe the tutors' availabilities are too
      restrictive, or there isn't enough flexibility to schedule a supertutor on the first day of workshops.
      Note: the constraints appear after the line "# --------- The constraints ---------", and look like
      variable = {i: m.addConstr( ... ) ...} etc. To get rid of one constraint, comment out everything from the
      "variable = {" down to the closing brace "}"
