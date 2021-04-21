# DATS-shift-signup

This program iterates through all operator and calls AssignRun function to determine a shift to be assigned to an operator for a big signup.
Version 1.1
Written by Jesse Xi Chen

Pre-conditions:
1st sheet of the file contain main information about the signing times for each operator
  Program assumes the list is sorted in ascending seniority order
  2nd column list the badge numbers of the operator
  4th column may contain 'not signing'
Digital sign up choices form is imported as the 3rd sheet. 

Post-conditions:
Choices assigned to operator is updated in the 5th column of the main sheet.

TODO:
For Ver 1.2
-Create valid input list and compare input against it.
-Testing for robustness, bad input.
-Optimize functions.

For Ver 1.3
-Allow continuation of sign up process in the middle of the list. Event trigger upon new response from the google form doucment (shift choice form).

