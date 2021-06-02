#!/usr/bin/env python

'''Here we modify an excel file that has a 'Time' column with improper values.
The values are replaced with a time that is incremented by a specified period.
The time elapsed has the requirement of having 6 decimal places.
datetime.datetime object is used because it can be incremented with
datetime.timedelta''' 

__author__ = 'Amir Farhadi'

import pandas as pd
import datetime

elapsed_time = 0
elapsed_time_list = []

file = 'Test.xlsx'
file_name_suffix = '_fixed' #new file has this is appended to its name
file_name = file.split('.')[0] #filename without extension
file_extension = file.split('.')[1] #file extension only

new_file = f'{file_name}{file_name_suffix}.{file_extension}' #name for the newly
                                                             #generated file

period = datetime.timedelta(seconds = 5)

df = pd.read_excel(file) # read file into dataframe

time_column = df['Time']

start_time = time_column[0] # first value in 'Time' column

start_time = start_time.split('.') # seperate the values

start_time = [int(element) for element in start_time] # convert them all to int

start_time[3] *= 1000 # convert the milliseconds to microseconds to simplify
                      # its use in the datetime object

# 1111, 1, 1, is just a dummy date since it is not needed for our purposes but
# is needed for the datetime object
time = datetime.datetime(1111, 1, 1, start_time[0], start_time[1],
                         start_time[2], start_time[3])

# Fix 'Time' column and create list for elapsed time
for i in range(len(time_column)):
    elapsed_time_6dec = f'{elapsed_time:.6f}' # give it 6 decimal places
    elapsed_time_list.append(elapsed_time_6dec.replace('.', ','))
    time_column[i] = f'{time.strftime("%H")}.' + \
                     f'{time.strftime("%M")}.' + \
                     f'{time.strftime("%S")}.' + \
                     f'{time.strftime("%f")[:-3]}' #get the millisec by cutting off last
                                                   #last 3 digits which are 000
    time += period # add 5 seconds to the time
    elapsed_time += period.total_seconds()

# Add Elapsed time(s) column
df.insert(2, "Elapsed time(s)", elapsed_time_list, True)

# Write to excel file
with pd.ExcelWriter(new_file) as writer:
    df.to_excel(writer, index = False)
