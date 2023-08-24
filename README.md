**Excel_Report_Clean_1**
A python program that takes an Excel or CSV based input file, cleans, and filters the data, then exports to a new file with multiple tabs based unique values. 
•	Allows user input for selecting file path, perfect for creating a sharable executable. 
  o	Required Imports:
    o	import pandas as pd
    o	import PySimpleGUI as sg
    o	import os
    o	import datetime
    o	import csv
    o	from datetime import datetime
    o	import openpyxl
•	This program was created with the intent of being semi-modular and dynamic. I set this up with the intention of automating other departments excel file cleaning. 
  o	Use cases: 
    o	Big reports where a lot of rows get dropped 
    o	If you want unique data separated onto different tabs in an Excel output
    o	If there is a list of stings that can change
    o	Able to make an executable that the end user can run to fine and clean their report
    o	Functions can be used all together or separate
•	Usually if you're decent at Excel this stuff doesn't take that long, but as many of us know a lot of people really aren't that good at Excel. 
•	Have been able to shave off about 10 hours a week in data cleaning with this program.
