import pandas as pd
import numpy as np
import re
import argparse


def natural_sort_key_vertical(well_position):
    """
    Extract row letter and column number for vertical sorting (A1, A2, ..., A12, B1, B2, ...)
    """
    match = re.match(r'([A-Z]+)(\d+)', well_position)
    if match:
        return (match.group(1), int(match.group(2)))
    return (well_position, 0)


def natural_sort_key_horizontal(well_position):
    """
    Extract row letter and column number for horizontal sorting (A1, B1, C1, ..., A2, B2, C2, ...)
    """
    match = re.match(r'([A-Z]+)(\d+)', well_position)
    if match:
        return (int(match.group(2)), match.group(1))  # Column first, then row
    return (0, well_position)


def process_sheet(df, sheet_name, sort_key_func):
    """
    Process a sheet by grouping by 'Well Position' and extracting 'Rn' values.
    
    Parameters:
    -----------
    df : DataFrame
        The input dataframe with 'Well Position' and 'Rn' columns
    sheet_name : str
        Name of the sheet (for logging)
    sort_key_func : callable
        Function to use for sorting well positions
    
    Returns:
    --------
    DataFrame with Time column and well positions as columns.
    """
    # Group the DataFrame by the positions column and aggregate the values into lists
    grouped_df = df.groupby('Well Position')['Rn'].apply(list).reset_index()
    
    # Set the indexed positions as the new index of the DataFrame
    grouped_df.set_index('Well Position', inplace=True)
    
    # Sort the index using the provided sort function
    grouped_df = grouped_df.sort_index(key=lambda x: [sort_key_func(pos) for pos in x])
    
    table_data = {}
    
    # Iterate over each position index in sorted order
    for position_index in grouped_df.index:
        # Get the values corresponding to the current position index
        values_at_index = grouped_df.loc[position_index, 'Rn']
        
        # Add the values to the dictionary with the position index as the key
        table_data[position_index] = values_at_index
    
    # Pad all values to the same length
    max_length = max(len(v) for v in table_data.values())
    for key in table_data:
        table_data[key] += [np.nan] * (max_length - len(table_data[key]))
    
    # Create time values (assuming max_length cycles)
    time_values = pd.DataFrame({'Time': range(1, max_length + 1)})
    
    # Create a DataFrame from the dictionary
    table_df = pd.DataFrame(table_data)
    
    # Drop columns with NaN values
    table_df = table_df.dropna(axis=1)
    
    # Concatenate time values with the table
    result_df = pd.concat([time_values, table_df], axis=1)
    
    return result_df


def pre_process_rox_vertical(input_file, experiment_info):
    """
    Process SYBR and ROX sheets from raw Excel file with VERTICAL sorting (A1, A2, ..., A12, B1, B2, ...).
    
    Parameters:
    -----------
    input_file : str
        Path to the raw Excel file containing SYBR and ROX sheets
    experiment_info : dict
        Dictionary containing experiment metadata:
        - 'Date': experiment date
        - 'Type of Experiment': list with experiment type
        - 'Experiment Number': list with experiment number
        - 'Pathogen': list with pathogen name
        - 'Condition': list with condition description
    
    Returns:
    --------
    tuple : (sybr_df, rox_df) processed DataFrames
    Saves output Excel file with both sheets
    """
    # Read both sheets
    print("Reading SYBR sheet...")
    df_sybr = pd.read_excel(input_file, sheet_name='SYBR')
    print(f"SYBR shape: {df_sybr.shape}")
    print(f"SYBR columns: {df_sybr.columns.tolist()}")

    print("\nReading ROX sheet...")
    df_rox = pd.read_excel(input_file, sheet_name='ROX')
    print(f"ROX shape: {df_rox.shape}")
    print(f"ROX columns: {df_rox.columns.tolist()}")

    # Process both sheets with vertical sorting
    print("\nProcessing SYBR sheet (vertical)...")
    sybr_processed = process_sheet(df_sybr, 'SYBR', natural_sort_key_vertical)
    print(f"SYBR processed shape: {sybr_processed.shape}")

    print("Processing ROX sheet (vertical)...")
    rox_processed = process_sheet(df_rox, 'ROX', natural_sort_key_vertical)
    print(f"ROX processed shape: {rox_processed.shape}")

    # Create output filename
    output_filename = f"{experiment_info['Date']}_{experiment_info['Type of Experiment'][0]}_{experiment_info['Experiment Number'][0]}_{experiment_info['Pathogen'][0]}_{experiment_info['Condition'][0]}_clean.xlsx"

    # Write to Excel with both sheets
    print(f"\nSaving to {output_filename}...")
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        sybr_processed.to_excel(writer, sheet_name='SYBR', index=False)
        rox_processed.to_excel(writer, sheet_name='ROX', index=False)

    print(f"✓ Successfully saved {output_filename}")
    
    return sybr_processed, rox_processed


def pre_process_rox_horizontal(input_file, experiment_info):
    """
    Process SYBR and ROX sheets from raw Excel file with HORIZONTAL sorting (A1, B1, C1, ..., A2, B2, C2, ...).
    
    Parameters:
    -----------
    input_file : str
        Path to the raw Excel file containing SYBR and ROX sheets
    experiment_info : dict
        Dictionary containing experiment metadata:
        - 'Date': experiment date
        - 'Type of Experiment': list with experiment type
        - 'Experiment Number': list with experiment number
        - 'Pathogen': list with pathogen name
        - 'Condition': list with condition description
    
    Returns:
    --------
    tuple : (sybr_df, rox_df) processed DataFrames
    Saves output Excel file with both sheets
    """
    # Read both sheets
    print("Reading SYBR sheet...")
    df_sybr = pd.read_excel(input_file, sheet_name='SYBR')
    print(f"SYBR shape: {df_sybr.shape}")
    print(f"SYBR columns: {df_sybr.columns.tolist()}")

    print("\nReading ROX sheet...")
    df_rox = pd.read_excel(input_file, sheet_name='ROX')
    print(f"ROX shape: {df_rox.shape}")
    print(f"ROX columns: {df_rox.columns.tolist()}")

    # Process both sheets with horizontal sorting
    print("\nProcessing SYBR sheet (horizontal)...")
    sybr_processed = process_sheet(df_sybr, 'SYBR', natural_sort_key_horizontal)
    print(f"SYBR processed shape: {sybr_processed.shape}")

    print("Processing ROX sheet (horizontal)...")
    rox_processed = process_sheet(df_rox, 'ROX', natural_sort_key_horizontal)
    print(f"ROX processed shape: {rox_processed.shape}")

    # Create output filename with 'horizontal' suffix
    output_filename = f"{experiment_info['Date']}_{experiment_info['Type of Experiment'][0]}_{experiment_info['Experiment Number'][0]}_{experiment_info['Pathogen'][0]}_{experiment_info['Condition'][0]}_horizontal_clean.xlsx"

    # Write to Excel with both sheets
    print(f"\nSaving to {output_filename}...")
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        sybr_processed.to_excel(writer, sheet_name='SYBR', index=False)
        rox_processed.to_excel(writer, sheet_name='ROX', index=False)

    print(f"✓ Successfully saved {output_filename}")
    
    return sybr_processed, rox_processed


def average_rox(excel_file, sheet_name, minutes=5):
    """
    Calculate the average of the last N minutes for each well and save to the same Excel file.
    
    Parameters:
    -----------
    excel_file : str
        Path to the Excel file containing the ROX data
    sheet_name : str
        Name of the sheet to process (e.g., 'ROX', 'SYBR')
    minutes : int
        Number of minutes to average from the end (default: 5)
    
    Returns:
    --------
    DataFrame : Summary dataframe with well names and their averages, plus grand average
    """
    # Load the sheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    print(f"Processing {sheet_name} sheet from {excel_file}")
    print(f"Shape: {df.shape}")
    print(f"Total time points: {len(df)}")
    
    # Get the last N rows
    last_rows = df.tail(minutes)
    print(f"\nLast {minutes} time points: {last_rows['Time'].tolist()}")
    
    # Calculate the average of the last N minutes for each well (excluding Time column)
    well_columns = [col for col in df.columns if col != 'Time']
    averages = last_rows[well_columns].mean()
    
    # Create a summary DataFrame
    summary_df = pd.DataFrame({
        'Well': well_columns,
        'Average_Last5min': averages.values
    })
    
    print("\nAverage of Last 5 Minutes:")
    print(summary_df)
    
    # Calculate the grand average (average of all well averages)
    grand_average = averages.mean()
    print(f"\n{'='*50}")
    print(f"Grand Average (All Wells, Last {minutes} Min): {grand_average:.2f}")
    print(f"{'='*50}")
    
    # Create a summary sheet with the well averages and grand average
    summary_with_grand = summary_df.copy()
    summary_with_grand.loc[len(summary_with_grand)] = ['GRAND_AVERAGE', grand_average]
    
    # Save back to the same Excel file
    print(f"\nSaving summary to {excel_file}...")
    summary_sheet_name = f'Summary_Last{minutes}min_{sheet_name}'
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        summary_with_grand.to_excel(writer, sheet_name=summary_sheet_name, index=False)
    
    print(f"✓ Successfully saved! Added '{summary_sheet_name}' sheet with grand average.")
    
    return summary_with_grand


def rox_norm(excel_file, sybr_sheet='SYBR', rox_sheet='ROX', summary_sheet='Summary_Last5min_ROX'):
    """
    Normalize SYBR values using ROX as reference.
    Formula: SYBR_value × (ROX_grand_average) / (ROX_well_average)

    Parameters:
    -----------
    excel_file : str
        Path to the Excel file containing both SYBR and ROX data
    sybr_sheet : str
        Name of the SYBR sheet (default: 'SYBR')
    rox_sheet : str
        Name of the ROX sheet (default: 'ROX')
    summary_sheet : str
        Name of the ROX summary sheet with averages (default: 'Summary_Last5min_ROX')

    Returns:
    --------
    DataFrame : Normalized SYBR data
    """
    print("Loading SYBR data...")
    df_sybr = pd.read_excel(excel_file, sheet_name=sybr_sheet)
    print(f"SYBR shape: {df_sybr.shape}")

    print("Loading ROX summary...")
    df_rox_summary = pd.read_excel(excel_file, sheet_name=summary_sheet)
    print(f"ROX summary shape: {df_rox_summary.shape}")

    rox_averages = {}
    grand_average = None

    for _, row in df_rox_summary.iterrows():
        well = row['Well']
        avg = row['Average_Last5min']

        if well == 'GRAND_AVERAGE':
            grand_average = avg
        else:
            rox_averages[well] = avg

    if grand_average is None:
        raise ValueError(f"GRAND_AVERAGE not found in sheet {summary_sheet}")

    print(f"\nGrand Average (ROX): {grand_average:.2f}")
    print(f"Number of wells: {len(rox_averages)}")

    print("\nNormalizing SYBR values...")
    df_normalized = df_sybr.copy()

    for col in df_sybr.columns:
        if col == 'Time':
            continue

        if col in rox_averages:
            rox_avg = rox_averages[col]
            df_normalized[col] = df_sybr[col] * (grand_average / rox_avg)

    print(f"\nSaving normalized SYBR to {excel_file}...")
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_normalized.to_excel(writer, sheet_name='SYBR_norm', index=False)

    print("✓ Successfully saved normalized SYBR to 'SYBR_norm' sheet")
    print(f"\nNormalized SYBR shape: {df_normalized.shape}")

    return df_normalized


def sybr_final(excel_file, norm_sheet='SYBR_norm', output_sheet='final'):
    """
    Baseline-correct SYBR normalized data by subtracting each column's first value.

    For every non-Time column:
        corrected_value = value - first_value_of_that_column

    Parameters:
    -----------
    excel_file : str
        Path to the Excel file
    norm_sheet : str
        Sheet containing normalized SYBR data (default: 'SYBR_norm')
    output_sheet : str
        Destination sheet name for baseline-corrected data (default: 'final')

    Returns:
    --------
    DataFrame : Baseline-corrected SYBR data
    """
    print(f"Loading normalized SYBR from sheet '{norm_sheet}'...")
    df_norm = pd.read_excel(excel_file, sheet_name=norm_sheet)
    print(f"Normalized SYBR shape: {df_norm.shape}")

    df_final = df_norm.copy()

    # Subtract each well column by its first time-point value.
    for col in df_final.columns:
        if col == 'Time':
            continue
        first_value = df_final[col].iloc[0]
        df_final[col] = df_final[col] - first_value

    print(f"Saving baseline-corrected data to sheet '{output_sheet}'...")
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_final.to_excel(writer, sheet_name=output_sheet, index=False)

    print(f"✓ Successfully saved baseline-corrected data to '{output_sheet}' sheet")
    return df_final


def run_rox_pipeline(input_file, experiment_info, orientation='vertical', minutes=5):
    """
    Run the full workflow in one call:
    1) preprocess raw SYBR/ROX sheets,
    2) compute ROX last-N-minute averages,
    3) normalize SYBR using ROX grand/well averages,
    4) baseline-correct normalized SYBR into 'final' sheet.

    Parameters:
    -----------
    input_file : str
        Path to raw Excel file with SYBR and ROX sheets
    experiment_info : dict
        Metadata used to build output filename
    orientation : str
        'vertical' or 'horizontal' preprocessing orientation
    minutes : int
        Number of ending time points to average for ROX summary

    Returns:
    --------
    dict : {'output_file', 'sybr_processed', 'rox_processed', 'rox_summary', 'sybr_norm', 'final'}
    """
    orientation = orientation.lower().strip()
    if orientation not in ('vertical', 'horizontal'):
        raise ValueError("orientation must be 'vertical' or 'horizontal'")

    if orientation == 'vertical':
        sybr_processed, rox_processed = pre_process_rox_vertical(input_file, experiment_info)
        output_file = f"{experiment_info['Date']}_{experiment_info['Type of Experiment'][0]}_{experiment_info['Experiment Number'][0]}_{experiment_info['Pathogen'][0]}_{experiment_info['Condition'][0]}_clean.xlsx"
    else:
        sybr_processed, rox_processed = pre_process_rox_horizontal(input_file, experiment_info)
        output_file = f"{experiment_info['Date']}_{experiment_info['Type of Experiment'][0]}_{experiment_info['Experiment Number'][0]}_{experiment_info['Pathogen'][0]}_{experiment_info['Condition'][0]}_horizontal_clean.xlsx"

    print("\n" + "=" * 60)
    print(f"AVERAGING ROX (LAST {minutes} MINUTES)")
    print("=" * 60)
    rox_summary = average_rox(output_file, 'ROX', minutes=minutes)

    print("\n" + "=" * 60)
    print("NORMALIZING SYBR USING ROX")
    print("=" * 60)
    sybr_norm = rox_norm(
        output_file,
        sybr_sheet='SYBR',
        rox_sheet='ROX',
        summary_sheet=f'Summary_Last{minutes}min_ROX'
    )

    print("\n" + "=" * 60)
    print("BASELINE CORRECTION TO FINAL")
    print("=" * 60)
    final_df = sybr_final(output_file, norm_sheet='SYBR_norm', output_sheet='final')

    return {
        'output_file': output_file,
        'sybr_processed': sybr_processed,
        'rox_processed': rox_processed,
        'rox_summary': rox_summary,
        'sybr_norm': sybr_norm,
        'final': final_df,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Preprocess SYBR/ROX data and run ROX-based normalization pipeline")
    parser.add_argument(
        "--orientation",
        choices=["vertical", "horizontal"],
        default="vertical",
        help="Well sorting orientation for preprocessing (default: vertical)",
    )
    parser.add_argument(
        "--minutes",
        type=int,
        default=5,
        help="Number of last time points used for ROX averaging (default: 5)",
    )
    args = parser.parse_args()

    # File to process
    input_file = "Single Cas12 Pooled 4 GN gRNAs against GN with rox_20260429_231114_Multicomponent_20260504_155333.xlsx"

    # Define the additional information for saving the final version
    experiment_info = {
        'Date': '20260429',
        'Type of Experiment': ['CRISPR single'],
        'Experiment Number': ['1'],
        'Pathogen': ['GN'],
        'Condition': ['pool 4 gRNAs -GN - test']
    }

    # One-call end-to-end workflow: preprocess + average ROX + normalize SYBR
    results = run_rox_pipeline(
        input_file=input_file,
        experiment_info=experiment_info,
        orientation=args.orientation,
        minutes=args.minutes,
    )

    print("\nPipeline complete.")
    print(f"Output file: {results['output_file']}")
