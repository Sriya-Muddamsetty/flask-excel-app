import os
from flask import Flask, render_template, request, redirect, url_for
import pandas as pd

app = Flask(__name__)

# Path to the Excel file
EXCEL_FILE = r"data.xlsx"  # Use raw string for file path to avoid escape characters

# Create an empty Excel file if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=["ID"])  # Start with an "ID" column for unique row identification
    df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add_column', methods=['GET', 'POST'])
def add_column():
    if request.method == 'POST':
        column_name = request.form['column_name']
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        if column_name not in df.columns:
            df[column_name] = None  # Add the new column with empty values
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        return redirect(url_for('index'))
    return render_template('add_column.html')

@app.route('/add_data', methods=['GET', 'POST'])
def add_data():
    if request.method == 'POST':
        # Collect form data and prepare it for appending to the DataFrame
        data = {col: request.form[col] for col in request.form if col != 'submit'}
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

        # Create a DataFrame from the new data and append it to the existing DataFrame
        new_data = pd.DataFrame([data])  # Create a DataFrame from the new row
        df = pd.concat([df, new_data], ignore_index=True)  # Concatenate the new data to the original DataFrame

        # Save the updated DataFrame back to the Excel file
        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')

        return redirect(url_for('index'))
    
    # If it's a GET request, render the form with existing column names
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
    columns = df.columns.tolist()
    return render_template('add_data.html', columns=columns)

@app.route('/update_data', methods=['GET', 'POST'])
def update_data():
    if request.method == 'POST':
        row_id = request.form['row_id']
        column_name = request.form['column_name']
        new_value = request.form['new_value']
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        
        # Update the specified row and column with the new value
        df.loc[df['ID'] == row_id, column_name] = new_value
        
        # Save the updated DataFrame back to the Excel file
        df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        
        return redirect(url_for('index'))
    
    # If it's a GET request, retrieve rows and columns for selection
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
    rows = df['ID'].tolist()
    columns = df.columns.tolist()
    
    return render_template('update_data.html', rows=rows, columns=columns)

@app.route('/show_data', methods=['GET', 'POST'])
def show_data():
    if request.method == 'POST':
        # Retrieve the start and end row from the form (if filtering is needed)
        start_row = int(request.form['start_row'])
        end_row = int(request.form['end_row'])
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        
        # Select the rows based on start and end row indices
        selected_data = df.iloc[start_row:end_row + 1]
        
        # Render the selected data as an HTML table
        return render_template('show_data.html', data=selected_data.to_html(classes='table'))
    
    # If no POST request (GET request), just display all the data in the DataFrame
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

    # Render the entire DataFrame as an HTML table
    return render_template('show_data.html', data=df.to_html(classes='table'))

@app.route('/delete_data', methods=['GET', 'POST'])
def delete_data():
    if request.method == 'POST':
        row_id = request.form['row_id']
        
        try:
            # Read the Excel file with the 'openpyxl' engine
            df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
            
            # Ensure 'ID' column exists
            if 'ID' not in df.columns:
                return "Error: 'ID' column is missing in the file."
            
            # Convert row_id to string (or int, depending on your ID format)
            row_id = str(row_id)  # Convert to string if IDs are strings in the Excel sheet
            
            # Check if the row ID exists in the 'ID' column
            if row_id not in df['ID'].values:
                return f"Error: Row ID '{row_id}' not found."

            # Filter the DataFrame to exclude the row with the given ID
            df = df[df['ID'] != row_id]

            # Save the modified DataFrame back to the Excel file
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, index=False)

            return redirect(url_for('index'))

        except Exception as e:
            return f"An error occurred: {str(e)}"
    
    # If the method is GET, show the available rows to delete
    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
    rows = df['ID'].tolist()
    return render_template('delete_data.html', rows=rows)

@app.route('/delete_column', methods=['GET', 'POST'])
def delete_column():
    if request.method == 'POST':
        column_name = request.form['column_name']
        df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        if column_name in df.columns:
            df.drop(columns=[column_name], inplace=True)
            df.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        return redirect(url_for('index'))

    df = pd.read_excel(EXCEL_FILE, engine='openpyxl')
    columns = df.columns.tolist()
    return render_template('delete_column.html', columns=columns)

if __name__ == '__main__':
    app.run(debug=True)
