import pandas as pd
from astral.sun import sun
from astral import LocationInfo
from datetime import datetime
import pytz
import tkinter as tk
import os
from tkinter import filedialog, messagebox

DEFAULT_FILE = r"G:\Shared drives\O&M\NCC Automations\ShortfallAnalysis\Data\Simple Production Data Report_2025-09-18 10-09.xlsx"

# ------ Tkinter GUI -------
root = tk.Tk()
root.title("Production Analysis")

# --- GUI Setup ---
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill=tk.BOTH, expand=True)

# File Selection
file_frame = tk.Frame(main_frame)
file_frame.pack(fill=tk.X, pady=5)
file_label = tk.Label(file_frame, text="File Path:")
file_label.pack(side=tk.LEFT, padx=(0, 5))
file_path_var = tk.StringVar()
file_path_var.set(DEFAULT_FILE)
file_entry = tk.Entry(file_frame, textvariable=file_path_var, state='readonly', width=60)
file_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

def select_file():
    """Opens a file dialog to select an Excel file."""
    filepath = filedialog.askopenfilename(
        title="Select Production Data Report",
        filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
    )
    if filepath:
        file_path_var.set(filepath)

file_button = tk.Button(file_frame, text="Browse...", command=select_file)
file_button.pack(side=tk.LEFT, padx=(5, 0))

# Latitude and Longitude Inputs
coords_frame = tk.Frame(main_frame)
coords_frame.pack(fill=tk.X, pady=5)

# Create StringVars to hold the entry box content
lat_var = tk.StringVar()
lon_var = tk.StringVar()

# Set default values for the StringVars
lat_var.set("34.685558")
lon_var.set("-79.540157")

lat_label = tk.Label(coords_frame, text="Latitude:")
lat_label.pack(side=tk.LEFT)
lat_entry = tk.Entry(coords_frame, textvariable=lat_var, width=15)
lat_entry.pack(side=tk.LEFT, padx=5)

lon_label = tk.Label(coords_frame, text="Longitude:")
lon_label.pack(side=tk.LEFT, padx=(10, 0))
lon_entry = tk.Entry(coords_frame, textvariable=lon_var, width=15)
lon_entry.pack(side=tk.LEFT, padx=5)

# ------ END Tkinter GUI -------

def production_analysis(file, latitude, longitude):
    df = pd.read_excel(file, skiprows=4)
    print(df.columns)
    df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors= 'coerce')

    # -- Handle potential parsing errors and set timezone --

    # Drop rows where timestamp could not be parsed to a valid date
    df.dropna(subset=['Timestamp'], inplace=True)

    # Localize to US/Eastern timezone before performing time-based comparisons
    df['Timestamp'] = df['Timestamp'].dt.tz_localize('US/Eastern', ambiguous='infer')

    # 1. Define the observer's location (e.g., New York City)
    city = LocationInfo("Solar Site", "USA", "America/New_York", latitude, longitude)
    tz = pytz.timezone(city.timezone) # Get the pytz timezone object for localization

    # Define base sunrise/sunset times by month (hour, minute) for fallback.
    # These are illustrative values for your location (North Carolina) and should be
    # refined based on actual data or averages if precise fallbacks are critical.
    # Note: These times are assumed to be in the local timezone (America/New_York).
    base_times_by_month = {
        1: {'sunrise_h': 7, 'sunrise_m': 15, 'sunset_h': 17, 'sunset_m': 30}, # Jan
        2: {'sunrise_h': 6, 'sunrise_m': 50, 'sunset_h': 18, 'sunset_m': 0},  # Feb
        3: {'sunrise_h': 6, 'sunrise_m': 30, 'sunset_h': 19, 'sunset_m': 30}, # Mar (accounts for DST start)
        4: {'sunrise_h': 6, 'sunrise_m': 0, 'sunset_h': 19, 'sunset_m': 50},  # Apr
        5: {'sunrise_h': 5, 'sunrise_m': 40, 'sunset_h': 20, 'sunset_m': 15}, # May
        6: {'sunrise_h': 5, 'sunrise_m': 30, 'sunset_h': 20, 'sunset_m': 30}, # Jun
        7: {'sunrise_h': 5, 'sunrise_m': 45, 'sunset_h': 20, 'sunset_m': 20}, # Jul
        8: {'sunrise_h': 6, 'sunrise_m': 10, 'sunset_h': 19, 'sunset_m': 50}, # Aug
        9: {'sunrise_h': 6, 'sunrise_m': 35, 'sunset_h': 19, 'sunset_m': 10}, # Sep
        10: {'sunrise_h': 7, 'sunrise_m': 0, 'sunset_h': 18, 'sunset_m': 30}, # Oct
        11: {'sunrise_h': 6, 'sunrise_m': 30, 'sunset_h': 17, 'sunset_m': 0}, # Nov (accounts for DST end)
        12: {'sunrise_h': 7, 'sunrise_m': 0, 'sunset_m': 17, 'sunset_m': 15}, # Dec
    }

    # 2. Calculate sunrise and sunset for each unique day in the data
    # This is more efficient than calculating it for every single row
    unique_dates_df = pd.DataFrame({'date': df['Timestamp'].dt.date.unique()})
    sun_times_list = []

    for date_obj in unique_dates_df['date']:
        try:
            s = sun(city.observer, date=date_obj)
            sun_times_list.append({
                'date': date_obj,
                'sunrise': s['sunrise'],
                'sunset': s['sunset']
            })
        except ValueError as e:
            print(f"WARNING: Could not calculate sun times for {date_obj}: {e}. Using monthly base times as fallback.")
            month = date_obj.month
            base_info = base_times_by_month.get(month)

            if base_info:
                # Construct timezone-aware datetime objects for the fallback
                fallback_sunrise = tz.localize( # Use the imported datetime class directly
                    datetime(
                        date_obj.year, date_obj.month, date_obj.day,
                        base_info['sunrise_h'], base_info['sunrise_m'], 0
                    )
                )
                fallback_sunset = tz.localize( # Use the imported datetime class directly
                    datetime(
                        date_obj.year, date_obj.month, date_obj.day,
                        base_info['sunset_h'], base_info['sunset_m'], 0
                    )
                )
                sun_times_list.append({
                    'date': date_obj,
                    'sunrise': fallback_sunrise,
                    'sunset': fallback_sunset
                })
            else:
                print(f"ERROR: No base times defined for month {month}. Date {date_obj} will be excluded.")

    # Create a DataFrame with the successfully calculated sun times
    sun_times_df = pd.DataFrame(sun_times_list)

    # 3. Merge sun times back to the main DataFrame and filter
    # Add a temporary 'date' column for merging
    df['date'] = df['Timestamp'].dt.date

    # Merge the sun times based on the date
    df_merged = pd.merge(df, sun_times_df, on='date', how='inner')

    # Filter for daylight hours, but offset by one hour after sunrise and one hour before sunset
    sunlight_df = df_merged[
        (df_merged['Timestamp'] > (df_merged['sunrise'] + pd.Timedelta(hours=1))) &
        (df_merged['Timestamp'] < (df_merged['sunset'] - pd.Timedelta(hours=1)))
    ].drop(columns=['date', 'sunrise', 'sunset']) # Clean up helper columns

    print("Sunlight DataFrame:")
    print(sunlight_df)

    # 4. Filter for rows with potential underperformance (value <= 0)
    # Identify columns to check (all except 'Timestamp')
    columns_to_check = sunlight_df.columns.drop('Timestamp')

    # 5. Identify, group, and summarize underperformance events.
    # Melt the DataFrame to easily check all component values at once.
    melted_df = sunlight_df.melt(
        id_vars=['Timestamp'],
        value_vars=columns_to_check,
        var_name='Name', # This will be our component name column
        value_name='Value'
    )

    # Filter for only the underperforming entries (value <= 0)
    underperforming_entries = melted_df[melted_df['Value'] <= 0].sort_values(by=['Timestamp', 'Name'])

    # --- Filter out low-production inverter events ---
    # This step aims to reduce false positives for inverters on very cloudy days.
    # If an inverter reports <= 0, we check the average of the *other* inverters at that time.
    # If the average is also very low (e.g., < 5 kW), we assume it's due to conditions, not a fault.
    if not underperforming_entries.empty:
        inverter_columns = [col for col in columns_to_check if 'inverter' in col.lower()]
        indices_to_drop = []

        # Get all inverter events to iterate through them
        inverter_events = underperforming_entries[underperforming_entries['Name'].str.lower().str.contains('inverter', na=False)]
        for index, event in inverter_events.iterrows():
            event_timestamp = event['Timestamp']
            event_inverter_name = event['Name']

            # Get all other inverters at the same timestamp from the original sunlight_df
            other_inverters = sunlight_df[sunlight_df['Timestamp'] == event_timestamp]

            if not other_inverters.empty:
                # Select only the inverter columns, excluding the one that is underperforming
                other_inverter_names = [inv for inv in inverter_columns if inv != event_inverter_name]

                # Only proceed if there are other inverters to compare against
                if other_inverter_names:
                    # Get the Series of production values for the other inverters at that timestamp
                    other_inverter_productions = other_inverters[other_inverter_names].iloc[0]
                    
                    # Calculate the average production of the other inverters
                    avg_production = other_inverter_productions.mean()

                    # If the average is below the threshold, mark this event's index for removal
                    if avg_production < 10:
                        indices_to_drop.append(index)

        print('Index:' , indices_to_drop)
        underperforming_entries.drop(indices_to_drop, inplace=True)

    # --- Advanced Event Grouping to handle hourly data and overnight gaps ---
    if not underperforming_entries.empty:
        # Sort is crucial for this logic to work
        underperforming_entries = underperforming_entries.sort_values(by=['Name', 'Timestamp']).reset_index(drop=True)

        # Manually assign group numbers to correctly bridge overnight gaps
        event_group_counter = 0
        group_assignments = [-1] * len(underperforming_entries)
        group_assignments[0] = event_group_counter

        # Iterate through the rows to check for event breaks
        for i in range(1, len(underperforming_entries)):
            prev_row = underperforming_entries.loc[i-1]
            curr_row = underperforming_entries.loc[i]

            time_diff = curr_row['Timestamp'] - prev_row['Timestamp']

            # Condition for a new event:
            # 1. The component name changes.
            # 2. The time gap is more than 1 hour (i.e., not consecutive hourly data).
            #    AND the gap is NOT an overnight gap (e.g., ~15-16 hours between sunset and next sunrise).
            #    We check if the gap is outside the plausible overnight range (e.g., <14h or >18h).
            is_new_event = (
                (curr_row['Name'] != prev_row['Name']) or
                (time_diff > pd.Timedelta(hours=1) and not (pd.Timedelta(hours=14) < time_diff < pd.Timedelta(hours=18)))
            )

            if is_new_event:
                event_group_counter += 1
            group_assignments[i] = event_group_counter
        
        underperforming_entries['event_group'] = group_assignments

    # Now, group by these event blocks and aggregate to get start/end times.
    outage_events_grouped = underperforming_entries.groupby('event_group').agg(
        Name=('Name', 'first'),
        Start_Timestamp=('Timestamp', 'min'),
        End_Timestamp=('Timestamp', 'max')
    ).reset_index(drop=True)

    # --- Filter out events that overlap with a 'Meter' event ---
    meter_events = outage_events_grouped[outage_events_grouped['Name'].str.contains('Meter', case=False, na=False)].copy()

    if not meter_events.empty:
        # Get all non-meter events to filter them
        non_meter_events = outage_events_grouped[~outage_events_grouped['Name'].str.contains('Meter', case=False, na=False)].copy()
        indices_to_drop = []

        for _, meter_event in meter_events.iterrows():
            # Define the meter event's time window with a 1-hour buffer on each side.
            meter_start_buffer = meter_event['Start_Timestamp'] - pd.Timedelta(hours=1)
            meter_end_buffer = meter_event['End_Timestamp'] + pd.Timedelta(hours=1)
            # Find non-meter events that overlap with the current meter event's buffered window.
            # Overlap condition: (StartA <= EndB) and (EndA >= StartB)
            overlapping_indices = non_meter_events[
                (non_meter_events['Start_Timestamp'] <= meter_end_buffer) &
                (non_meter_events['End_Timestamp'] >= meter_start_buffer)
            ].index
            indices_to_drop.extend(overlapping_indices)
        
        # Drop the identified overlapping events from the main dataframe
        outage_events_grouped.drop(index=list(set(indices_to_drop)), inplace=True)

    if outage_events_grouped.empty:
        print("\nNo underperformance events found during sunlight hours.")
        return None # Return None to indicate no file was created

    # Create the 'Event_Number' column, which counts events per component.
    outage_events_grouped['Event_Number'] = outage_events_grouped.groupby('Name').cumcount() + 1
    outage_events_grouped['False Alarm?'] = '' #Column for User to insert Checkbutton.
    outage_events_grouped['Our Fault?'] = '' #Column for User to insert Checkbutton.
    # Reorder columns to the desired format
    final_outage_summary = outage_events_grouped[['Name', 'Event_Number', 'Start_Timestamp', 'End_Timestamp', 'False Alarm?', 'Our Fault?']].sort_values(by='Start_Timestamp')

    print("\nSummary of Outage Events:")
    print(final_outage_summary)
    
    # 6. Save the summary to an Excel file
    try:
        # Get the directory of the input file
        output_dir = os.path.dirname(file)
        # Create a timestamp string for a unique filename
        timestamp_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        output_filename = f"Shorthorn Outage Events_{timestamp_str}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        # Before saving to Excel, convert timezone-aware datetimes to naive ones.
        # This prevents issues where Excel might not display the timestamp correctly.
        final_outage_summary['Start_Timestamp'] = final_outage_summary['Start_Timestamp'].dt.tz_localize(None)
        final_outage_summary['End_Timestamp'] = final_outage_summary['End_Timestamp'].dt.tz_localize(None)

        # Save the DataFrame to Excel
        final_outage_summary.to_excel(output_path, index=False, sheet_name='Outage Events')

        print(f"\nSuccessfully saved outage summary to:\n{output_path}")
        return output_path # Return the path on success
    except Exception as e:
        print(f"\nERROR: Failed to save Excel file. Reason: {e}")
        return None # Return None on failure

def run_analysis():
    """Get values from GUI and run the production analysis."""
    file_path = file_path_var.get()
    try:
        latitude = float(lat_var.get())
        longitude = float(lon_var.get())
    except ValueError:
        messagebox.showerror("Invalid Input", "Latitude and Longitude must be valid numbers.")
        return

    if not file_path:
        messagebox.showerror("Invalid Input", "Please select a file to analyze.")
        return

    try:
        output_file_path = production_analysis(file_path, latitude, longitude)
        if output_file_path:
            messagebox.showinfo("Success", f"Analysis complete!\n\nOutput saved to:\n{output_file_path}")
        else:
            messagebox.showinfo("Analysis Complete", "No underperformance events were found to report.")
    except Exception as e:
        messagebox.showerror("Analysis Error", f"An error occurred during analysis:\n{e}")

analyze_button = tk.Button(main_frame, text="Run Analysis", command=run_analysis, bg="lightblue", font=('Arial', 10, 'bold'))
analyze_button.pack(pady=20)

root.mainloop()