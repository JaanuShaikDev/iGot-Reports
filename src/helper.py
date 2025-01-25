import pandas as pd
import os
from openpyxl import Workbook
import plotly.graph_objects as go


def get_untrained_data(course_names, data, grouped_data):

  '''This function will generate %data and trng_data'''

  # Create a dictionary to store the dataframes for each sheet
  perc_data = {}

  trng_data = {sub : {} for sub in (data[grouped_data].unique())}

  # Process each sheet
  for sheet_name in course_names[:-1]:  # Exclude the last sheet if it's not needed
      trained = pd.read_excel("Data/completed.xlsx", sheet_name=sheet_name, header=1)
      filter = data[(data['Employee No.'].isin(trained['Employee No.']))]
      total = data.groupby(grouped_data)[grouped_data].count()
      data_total = pd.DataFrame(
         {grouped_data: sorted(
            list(data[grouped_data].unique())), 'Total': list(total)}
         )
      untrained = filter.groupby(grouped_data)[grouped_data].count()
      data_trained = pd.DataFrame(
         {grouped_data: sorted(list(
            filter[grouped_data].unique())), 'Untrained': list(untrained)}
         )

      # Store the %_dataframe in the dictionary
      perc_data[sheet_name] = pd.merge(data_total, data_trained, on=grouped_data, how='left')
      perc_data[sheet_name]['%_of_completion'] = (
         ((perc_data[sheet_name]['Total'] - perc_data[sheet_name]['Untrained']) 
                                                   / perc_data[sheet_name]['Total']) * 100
                                                   ).round(2)

      # Store the untrained officials in the dictionary
      for sub in filter[grouped_data].unique():
          trng_data[sub][sheet_name] = filter[filter[grouped_data] == sub]

  return perc_data, trng_data


def grouped_files(data, trng_data, my_dir, grouped_data):

  ''' This function will generate group_wise Excel files'''

  for sub in list(data[grouped_data].unique()):
    file_path = os.path.join(my_dir, f'{sub}.xlsx')
    with pd.ExcelWriter(file_path, engine = "openpyxl") as writer:
       if not trng_data[sub]:
        writer.book = Workbook()
        writer.book.create_sheet("Sheet1")
       else:
        for sheet_name, df in trng_data[sub].items():
          df.to_excel(writer, sheet_name = f"{sheet_name}", index = False)

    #files.download(file_path)


def calc_perc(perc_data, my_dir):

  '''This function will write %_dataframes to Excel file'''

  file_path = os.path.join(my_dir, "Required.xlsx")
  with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
   for sheet_name, df in perc_data.items():
          df.to_excel(writer, sheet_name=sheet_name, index=False)

  #files.download(file_path)


def course_file_path_data():

  '''This function will give data file, course_names and file_path to download files'''

  er = pd.read_excel("Data/er.xlsx")
  sd = pd.read_excel("Data/sd.xlsx")
  sol = pd.read_excel("Data/sol.xlsx")

  #er = er[er['Employee No.'] != 0]
  er.drop_duplicates(subset=['Employee No.'], inplace = True)
  #er.reset_index()
  #er.drop(columns=['index'], axis = 1, inplace = True)
  er_new = pd.merge(er, sd, left_on = 'Facility Id',
                     right_on = 'FACILITY_ID', how = 'left', indicator = True)

  cols = ['Facility Id','Facility Description','Employee No.',
          'Employee Name', 'Cadre', 'SOL_ID', 'PROFIT_CENTRE', 'SUB_DIVISION']

  data = er_new[cols].merge(
     sol, left_on = 'SOL_ID', right_on = 'SOL ID', how = 'left'
     )
  data.drop(
     columns = ['SOL_ID', 'SOL ID', 'Facility Id', 'PROFIT_CENTRE'], axis = 1, inplace = True
     )
  data.rename(
     columns = {'Facility Description':'Office', 'Office Name' : 'Account Office',
                 'SUB_DIVISION' : 'Sub Division' } , inplace = True
                 )
  data = data[['Employee No.', 'Employee Name', 'Cadre', 'Office', 'Account Office', 'Sub Division']]

  my_dir = "Reports"
  os.makedirs(my_dir, exist_ok=True)

  # Get sheet names from the Excel file
  temp = pd.ExcelFile("Data/completed.xlsx")
  course_names = temp.sheet_names

  return data, course_names, my_dir


def plot_pichart(perc_data):

  '''This function will plot Pi-Chart for overall training of divsion'''

  # import plotly.graph_objects as go

  # Initialize variables to aggregate overall data
  total_trained = 0
  total_untrained = 0
  total_count = 0

  # Process and aggregate data from all sheets
  for sheet_name, df in perc_data.items():
      required = df.copy()
      required = required[required['Sub Division'] != 'Narasaraopet Division'].reset_index()
      required.drop(columns=['index'], inplace=True)

      # Calculate the overall totals for trained and untrained employees
      total_trained += required['Total'].sum() - required['Untrained'].sum()
      total_untrained += required['Untrained'].sum()
      total_count += required['Total'].sum()

  # Calculate percentage of completion
  completion_percentage = (total_trained / total_count) * 100
  untrained_percentage = (total_untrained / total_count) * 100

  # Data for Pie chart
  labels = ['Completed', 'Untrained']
  values = [completion_percentage, untrained_percentage]

  # Create pie chart
  fig = go.Figure(data=[go.Pie(
      labels=labels,
      values=values,
      textinfo='label+percent',
      hoverinfo='label+percent+value',
      marker=dict(colors=['#00b300', '#ff6666']),
  )])

  # Update layout
  fig.update_layout(
      title="Completion Percentage for Narasaraopet Division in iGOT Courses (Dec-2024)",
      height=600,
      width=800,
      showlegend=True,
      uniformtext_minsize=10,
      uniformtext_mode='hide'
  )

  # Show the pie chart
  fig.show()


def plot_sd_wise_data(perc_data):

  '''This function will plot overall % of completion of training by sub divisions'''

  # import plotly.graph_objects as go

  # Initialize variables for overall completion data by Sub Division
  subdivision_data = {}

  # Process each sheet to calculate training completion percentage by Sub Division
  for sheet_name, df in perc_data.items():
      required = df.copy()
      required = required[required['Sub Division'] != 'Narasaraopet Division'].reset_index()
      required.drop(columns=['index'], inplace=True)

      # Aggregate data for each Sub Division
      for sub_div in required['Sub Division'].unique():
          sub_data = required[required['Sub Division'] == sub_div]
          total = sub_data['Total'].sum()
          untrained = sub_data['Untrained'].sum()
          trained = total - untrained

          # Store the data for each Sub Division
          if sub_div not in subdivision_data:
              subdivision_data[sub_div] = {'trained': 0, 'total': 0}

          subdivision_data[sub_div]['trained'] += trained
          subdivision_data[sub_div]['total'] += total

  # Prepare data for bar chart
  sub_divisions = []
  trained_percentages = []

  for sub_div, data in subdivision_data.items():
      sub_divisions.append(sub_div)
      percentage_trained = (data['trained'] / data['total']) * 100 if data['total'] > 0 else 0
      trained_percentages.append(percentage_trained)

  # Create bar chart
  fig = go.Figure()
  fig.add_trace(go.Bar(
      x=sub_divisions,
      y=trained_percentages,
      name='Percentage Trained',
      marker=dict(color=required["%_of_completion"], colorscale="Viridis")
  ))

  # Update layout for the bar chart
  fig.update_layout(
      title="Overall Percentage Completion by Sub Division for iGOT Courses (Dec-2024)",
      height=600,
      width=800,
      xaxis_title="Sub Division",
      yaxis_title="Percentage of Completion",
      yaxis=dict(ticksuffix='%'),
      legend_title="Category"
  )

  # Show the bar chart
  fig.show()


def count_employee_occurrences(trng_data, data, grouped_data, column_name='Employee No.'):
    """
    This function counts the occurrences of each 'Employee No.' across all sheets in the trng_data dictionary.

    :param trng_data: Dictionary of dataframes representing different sheets
    :param column_name: The column (default 'Employee No.') whose occurrences you want to count
    """
    # Initialize an empty list to store counts from all sheets
    all_counts = pd.Series(dtype=int)

    # Iterate through each sheet in the trng_data dictionary
    for sub, sheet_data in trng_data.items():
        for sheet_name, df in sheet_data.items():
            if column_name in df.columns:
                # Count occurrences of 'Employee No.' in the current sheet
                sheet_counts = df[column_name].value_counts()
                # Add the counts from the current sheet to the total counts
                all_counts = all_counts.add(sheet_counts, fill_value=0)
            else:
                print(f"Column '{column_name}' not found in sheet {sheet_name}")

    # Convert the Series to DataFrame
    employee_counts_df = all_counts.reset_index()
    employee_counts_df.columns = ['Employee No.', 'Pending_Count']
    employee_counts_df = pd.merge(employee_counts_df, data, on = 'Employee No.', how = 'right')
    employee_counts_df = employee_counts_df[~employee_counts_df['Pending_Count'].isna()].reset_index()
    employee_counts_df.drop(columns=['index'], axis = 1, inplace = True)
    employee_counts_df = employee_counts_df[
       ['Employee No.', 'Employee Name', 'Cadre', 'Office', 'Account Office', 'Sub Division', 'Pending_Count']
       ]


    for sub in list(data[grouped_data].unique()):
      file_path = os.path.join(my_dir, f'{sub}.xlsx')
      #employee_counts_df[(employee_counts_df[grouped_data] == sub) & ((employee_counts_df['Pending_Count'] > 5) & (employee_counts_df['Pending_Count'] < 10))].to_excel(file_path, index = False)
      employee_counts_df[
         (employee_counts_df[grouped_data] == sub) & (employee_counts_df['Pending_Count']  > 0)
         ].to_excel(file_path, index = False)

      #files.download(file_path)
