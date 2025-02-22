from flask import Flask, request, render_template, send_file
import csv
from tabulate import tabulate
from collections import defaultdict
from openpyxl import Workbook
import io
import os
import base64  # Add this import

app = Flask(__name__)


def safe_str(val):
    """Convert value to string safely, handling lists and None."""
    if isinstance(val, list):
        return safe_str(val[0]) if val else ''
    if val is None:
        return ''
    return str(val).strip()


def try_decode_file(file_content, encodings=['utf-8-sig', 'latin1', 'windows-1252', 'iso-8859-1']):
    """Attempt to decode file content with multiple encodings."""
    for encoding in encodings:
        try:
            return file_content.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("Unable to decode file with available encodings")


def read_csv_parts(csv_content):
    """Read part info including Weight(kg/m) from CSV content."""
    parts = []
    try:
        reader = csv.DictReader(csv_content.splitlines(), delimiter=';')
        for row in reader:
            if not any(row.values()):
                continue
            row_stripped = {safe_str(k): safe_str(v) for k, v in row.items()}
            profile = f"{row_stripped.get('Size', '')}_{row_stripped.get('Grade', '')}"
            if not profile or profile == '_':
                continue
            try:
                length_val = int(row_stripped.get('Length(mm)', 0))
                demand_val = int(row_stripped.get('Quantity', 0))
                weight_per_m = float(row_stripped.get('Weight(kg/m)', 0))
            except ValueError:
                continue
            if length_val <= 0 or demand_val <= 0:
                continue
            parts.append((profile, length_val, demand_val, weight_per_m))
        return parts
    except Exception:
        return None


def best_fit_cutting_stock(parts, stock_length, cut_kerf=0.0):
    """Best-Fit cutting algorithm."""
    items = []
    for profile, length, demand, _ in parts:
        items.extend([(profile, length)] * demand)
    items.sort(key=lambda x: x[1], reverse=True)

    bins = []
    for profile, length in items:
        effective_length = length + cut_kerf
        best_bin = None
        min_remaining = float('inf')

        for bin in bins:
            if bin['remaining'] >= effective_length:
                remaining_after = bin['remaining'] - effective_length
                if remaining_after < min_remaining:
                    min_remaining = remaining_after
                    best_bin = bin

        if best_bin:
            best_bin['remaining'] -= effective_length
            best_bin['cuts'].append((profile, length))
        else:
            bins.append({
                'remaining': stock_length - effective_length,
                'cuts': [(profile, length)]
            })
    return bins


def generate_pattern_details_table(bins, stock_length):
    """Generate cutting pattern details."""
    headers = ['Pattern Name', 'Pattern Length (mm)', 'Cut Details', 'Remaining Waste (mm)']
    table_data = []
    for i, bin in enumerate(bins, 1):
        pattern_name = f"Pattern {i}"
        total_length = sum(length for _, length in bin['cuts'])
        cut_details = ' + '.join(f"1x {profile}({length}mm)" for profile, length in bin['cuts'])
        remaining_waste = bin['remaining']
        pattern_length = stock_length - remaining_waste
        table_data.append([pattern_name, pattern_length, cut_details, remaining_waste])
    return table_data, headers


def generate_final_report(parts, bins, stock_length, cut_kerf=0.0, profile=None):
    """Generate summary report with specified formats and weight calculations."""
    total_stocks_used = len(bins)
    total_stock_consumed_mm = total_stocks_used * stock_length
    effective_usage_mm = sum(length * demand for _, length, demand, _ in parts)
    waste_mm = total_stock_consumed_mm - effective_usage_mm - cut_kerf * sum(demand for _, _, demand, _ in parts)

    total_stock_consumed_m = total_stock_consumed_mm / 1000
    effective_usage_m = effective_usage_mm / 1000
    waste_m = waste_mm / 1000

    weight_per_m = parts[0][3]
    total_order_weight_kg = total_stock_consumed_m * weight_per_m if weight_per_m > 0 else 0
    effective_weight_kg = effective_usage_m * weight_per_m if weight_per_m > 0 else 0
    waste_weight_kg = waste_m * weight_per_m if weight_per_m > 0 else 0

    headers = ['Profile', 'Total Stocks Used', 'Stock Length (mm)', 'Weight (kgm)',
               'Total Usage (m)', 'Effective Usage (m)', 'Waste (m)',
               'Total Order Weight (kg)', 'Effective Weight (kg)', 'Waste Weight (kg)']
    table_data = [[profile if profile else parts[0][0],
                   total_stocks_used, stock_length, weight_per_m,
                   total_stock_consumed_m, effective_usage_m, waste_m,
                   total_order_weight_kg, effective_weight_kg, waste_weight_kg]]
    return table_data, headers


@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Main route handler for the root URL ('/').
    Handles both GET requests (showing the form) and POST requests (processing the uploaded CSV).

    Methods:
        GET: Displays the input form (index.html)
        POST: Processes the CSV file, generates cutting patterns, and shows results
    """
    if request.method == 'POST':
        # Handle form submission when user uploads a CSV file

        # Check if a file was uploaded
        if 'csv_file' not in request.files:
            return render_template('index.html', error='No file uploaded')

        file = request.files['csv_file']
        # Verify that a file was selected
        if file.filename == '':
            return render_template('index.html', error='No file selected')

        # Get form inputs with type conversion
        stock_length = request.form.get('stock_length', type=int)
        cut_kerf = request.form.get('cut_kerf', default=0.0, type=float)

        # Validate stock length
        if not stock_length or stock_length <= 0:
            return render_template('index.html', error='Invalid stock length')

        # Get selected encoding from form, default to 'utf-8-sig'
        selected_encoding = request.form.get('encoding', 'utf-8-sig')
        encodings = [selected_encoding] + [e for e in ['utf-8-sig', 'latin1', 'windows-1252', 'iso-8859-1'] if
                                           e != selected_encoding]

        try:
            # Read the uploaded file as bytes
            file_content = file.read()
            # Try to decode with selected encodings
            csv_content, used_encoding = try_decode_file(file_content, encodings)
            # Parse CSV content into parts
            parts = read_csv_parts(csv_content)

            # Check if parsing succeeded
            if not parts:
                return render_template('index.html',
                                       error=f'Failed to parse CSV file (tried encoding: {used_encoding})')
        except UnicodeDecodeError as e:
            return render_template('index.html',
                                   error=f'File encoding error: {str(e)}. Try saving the CSV with UTF-8 encoding.')
        except Exception as e:
            return render_template('index.html',
                                   error=f'Error processing file: {str(e)}')

        # Group parts by profile
        parts_by_profile = defaultdict(list)
        for part in parts:
            parts_by_profile[part[0]].append(part)

        # Initialize lists for storing results
        all_final_data = []
        all_details_data = []
        final_headers = None
        details_headers = None

        # Process each profile group
        for profile, group in parts_by_profile.items():
            # Generate cutting patterns using best-fit algorithm
            bins = best_fit_cutting_stock(group, stock_length, cut_kerf)
            # Generate final report for this profile
            agg_table, agg_headers = generate_final_report(group, bins, stock_length, cut_kerf, profile)
            if agg_table:
                all_final_data.extend(agg_table)
            if not final_headers:
                final_headers = agg_headers

            # Generate pattern details table
            details_table, det_headers = generate_pattern_details_table(bins, stock_length)
            if details_table:
                all_details_data.append([f"Profile: {profile}", '', '', ''])
                all_details_data.extend(details_table)
            if not details_headers:
                details_headers = det_headers

        # Generate Excel file in memory
        wb = Workbook()
        ws = wb.active
        ws.title = "Cutting Stock Results"

        # Write Final Aggregate Report to Excel
        ws.append(["--- Final Aggregate Report ---"])
        ws.append(final_headers)
        for row in all_final_data:
            ws.append(row)

        # Add spacing and Pattern Details
        ws.append([])
        ws.append(["--- Pattern Details Table ---"])
        ws.append(details_headers)
        for row in all_details_data:
            ws.append(row)

        # Save Excel to memory buffer
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        # Convert Excel bytes to base64 string for safe HTML transmission
        excel_data = base64.b64encode(excel_buffer.getvalue()).decode('utf-8')

        # Render the results template with all data
        return render_template('results.html',
                               final_data=all_final_data,  # Final report table data
                               final_headers=final_headers,  # Final report headers
                               details_data=all_details_data,  # Pattern details table data
                               details_headers=details_headers,  # Pattern details headers
                               excel_data=excel_data)  # Base64-encoded Excel file

    # For GET requests, show the input form
    return render_template('index.html')


@app.route('/download_excel')
def download_excel():
    """
    Route handler for downloading the generated Excel file.
    Takes the base64-encoded Excel data from the query parameter and sends it as a file.
    """
    # Get the base64-encoded Excel data from URL parameters
    excel_data = request.args.get('excel_data')

    # Decode base64 string back to bytes
    buffer = io.BytesIO(base64.b64decode(excel_data))

    # Send the file to the user as a download
    return send_file(buffer,
                     download_name='cutting_stock_results.xlsx',  # File name for download
                     as_attachment=True,  # Force download instead of display
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')  # Excel MIME type


if __name__ == '__main__':
    """
    Entry point for running the Flask application.
    Configures the app to run on all interfaces and uses an environment-specified port.
    """
    # Get port from environment variable (useful for hosting platforms) or default to 5000
    port = int(os.environ.get('PORT', 5000))

    # Run the app, binding to all network interfaces (0.0.0.0)
    # This makes it accessible externally when deployed
    app.run(host='0.0.0.0', port=port)