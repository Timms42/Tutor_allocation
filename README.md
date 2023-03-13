# Tutor_allocation
Program and mathematical model for optimal allocation of university tutors to course tutorials.

scie_tutor_alloc.py implements a mixed-integer programming model for optimal tutor-workshop allocations, using Gurobi as a solver.
It reads in data from a provided Excel spreadsheet, and is set up so the user is not required to understand or interact with the underlying mathematical model.

Tutor_allocation_model describes the mixed-integer mathematical program that underpins the allocation code.
