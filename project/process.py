import pandas as pd
from datetime import datetime
from project.utils import *

def work_summary(df):

    df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d')
    df['Time'] = pd.to_datetime(df['Time'], format='%H:%M:%S').dt.time

    # Combine Date and Time into a single datetime column for easier calculation
    df['DateTime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str))

   
    grouped = df.groupby(['E. ID', df['Date'].dt.date])

    # Step 5: Calculate Working Hours
    working_times = []

    for (user_id, date), group in grouped:
        # Filter for 'I' and 'O' events
        logins = group[group['Status'] == 'I']['DateTime'].tolist()
        logouts = group[group['Status'] == 'O']['DateTime'].tolist()
        
        # Pair logins and logouts and calculate the working time
        total_work_time = pd.Timedelta(0)
        for login, logout in zip(logins, logouts):
            total_work_time += (logout - login)

        total_work_time_str = str(total_work_time).split(" ")[-1]
        
        working_times.append([user_id, date, total_work_time_str])

    # Create a DataFrame with the results
    result_df = pd.DataFrame(working_times, columns=['E. ID', 'Date', 'Total Working Time'])
    result_df['Date'] = result_df['Date'].apply(lambda x: str(x))

    # Display the results
    # print(result_df)

    # Optionally, save to a file
    # result_df.to_csv('working_times.txt', sep='\t', index=False)
    return result_df
    


def process_file(input_file):
    data = []
    with open(input_file, 'r') as file:
        for line in file:
            parts = line.strip().split(',')
            data.append(parts)

    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=['Status', 'E. ID', 'Date', 'Time'])
    modified_df = process_dataframe(df)
    
    result = work_summary(modified_df)

    # modified_df['Difference(O-I)'] = ''
    # try:
    #     for i in range(0, len(modified_df) - 1, 2):
            
    #         time_out = modified_df.at[i, 'Time']
    #         time_in = modified_df.at[i + 1, 'Time']
    #         diff = calculate_time_difference(time_out, time_in)
    #         modified_df.at[i, 'Difference(O-I)'] = str(diff)

    # except Exception as e:
    #     print(e)

    modified_df['Time'] = modified_df['Time'].apply(lambda x: convert_time_format(x))
    modified_df['Date'] = modified_df['Date'].apply(lambda x: str(x).split(' ')[0])
    modified_df = modified_df.drop("DateTime", axis=1)

    # modified_df.to_csv('processed_data.txt', sep='\t', index=False)

    return modified_df, result
