import pandas as pd
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Side, Border
import plotly.graph_objects as go
from src.helper import get_untrained_data, course_file_path_data, get_untrained_data
from src.helper import grouped_files, calc_perc, plot_sd_wise_data, count_employee_occurrences
from src.helper import plot_pichart, style_border


data, course_names = course_file_path_data()

perc_data, trng_data = get_untrained_data(course_names, data, 'Sub Division')
#perc_data, trng_data = get_untrained_data(course_names, data, 'Account Office')

grouped_files(data, trng_data, 'Sub Division')
#grouped_files(data, trng_data, my_dir, 'Account Office')

#calc_perc(perc_data, my_dir)

#plot_course_wise_data(perc_data) # Run for sub-division wise only

plot_sd_wise_data(perc_data) # Run for sub-division wise only

plot_pichart(perc_data) # Run for sub-division wise only

count_employee_occurrences(trng_data, data, 'Sub Division')
#count_employee_occurrences(trng_data, data, 'Account Office')

#style_border("Reports/Count")