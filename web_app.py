from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pandas as pd
import time
import requests
import os
import subprocess
import pyautogui
from functools import wraps

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Needed for flashing messages
app.config['UPLOAD_FOLDER'] = 'uploads/'

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

class GoogleTrendsAgent:
    def __init__(self, serp_api_key):
        self.api_key = serp_api_key
        self.base_url = "https://serpapi.com/search.json"

    def get_trends_data(self, keyword):
        params = {
            'engine': 'google_trends',
            'q': keyword,
            'data_type': 'TIMESERIES',
            'date': 'now 7-d',
            'geo': '',
            'api_key': self.api_key
        }
        try:
            response = requests.get(self.base_url, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()
            if 'error' in data:
                return []
            return data.get('interest_over_time', {}).get('timeline_data', [])
        except Exception:
            return []

    def get_region_data(self, keyword):
        params = {
            'engine': 'google_trends',
            'q': keyword,
            'data_type': 'GEO_MAP_0',
            'date': 'now 7-d',
            'api_key': self.api_key
        }
        try:
            response = requests.get(self.base_url, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()
            if 'error' in data:
                return []
            return data.get('interest_by_region', [])
        except Exception:
            return []

    def check_threshold_criteria(self, interest_values):
        if not interest_values:
            return False
        days_above_50 = sum(1 for value in interest_values if value and value > 50)
        return days_above_50 >= 2

    def check_country_criteria(self, region_data):
        target_countries = ['United States', 'Germany', 'Japan', 'United Kingdom', 'France', 'China']
        top_countries = [item.get('location', '') for item in region_data[:10]]
        return any(country in top_countries for country in target_countries)

    def extract_interest_values(self, trends_data, keyword):
        interest_values = []
        for item in trends_data:
            values = item.get('values', [])
            if values:
                if len(values) == 1:
                    interest_values.append(values[0].get('extracted_value', 0))
                else:
                    for value in values:
                        if value.get('query', '').lower() == keyword.lower():
                            interest_values.append(value.get('extracted_value', 0))
                            break
                    else:
                        interest_values.append(0)
        return interest_values

    def analyze_keyword(self, keyword):
        trends_data = self.get_trends_data(keyword)
        if not trends_data:
            return None
        interest_values = self.extract_interest_values(trends_data, keyword)
        if len(interest_values) < 7:
            return None
        time.sleep(1)
        region_data = self.get_region_data(keyword)
        meets_threshold = self.check_threshold_criteria(interest_values)
        has_target_countries = self.check_country_criteria(region_data)
        if meets_threshold and has_target_countries:
            return {
                'keyword': keyword,
                'interest_values': interest_values,
                'max_interest': max(interest_values) if interest_values else 0,
                'avg_interest': round(sum(interest_values) / len(interest_values)) if interest_values else 0,
                'days_above_50': sum(1 for v in interest_values if v and v > 50),
                'top_countries': [item.get('location', '') for item in region_data[:5]]
            }
        return None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        api_key = request.form.get('api_key')
        file = request.files.get('file')
        if not api_key or not file:
            flash('API key and file are required!')
            return redirect(request.url)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        return redirect(url_for('process_keywords', api_key=api_key, file_path=filepath))
    return render_template('index.html')

@app.route('/process_keywords')
def process_keywords():
    api_key = request.args.get('api_key')
    file_path = request.args.get('file_path')
    if not api_key or not file_path:
        return 'Missing API key or file path', 400
    try:
        df = pd.read_excel(file_path)
        required_columns = ['Keywords', 'AVG. Search', 'Competition', 'RID']
        if not all(col in df.columns for col in required_columns):
            return 'Missing required columns in Excel file', 400
        df['AVG. Search'] = pd.to_numeric(df['AVG. Search'], errors='coerce')
        filtered_df = df[(df['AVG. Search'] >= 5000) & (df['Competition'] == 'Low')]
        # Remove the word 'Market' from the end of each keyword (if present)
        filtered_df['Keywords'] = filtered_df['Keywords'].str.replace(r'\s*Market$', '', regex=True)
        keywords = filtered_df['Keywords'].dropna().tolist()
        agent = GoogleTrendsAgent(api_key)
        qualifying_keywords = []
        for keyword in keywords:
            result = agent.analyze_keyword(keyword)
            if result:
                qualifying_keywords.append(result)
        qualifying_keyword_list = [result['keyword'] for result in qualifying_keywords]
        qualifying_rid_df = filtered_df[filtered_df['Keywords'].isin(qualifying_keyword_list)]
        final_rids = qualifying_rid_df['RID'].tolist()
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'Qualifying_Keywords_RIDs.csv')
        pd.DataFrame({'RID': final_rids}).to_csv(output_file, index=False)
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f'Error processing file: {e}', 500

@app.route('/task', methods=['GET', 'POST'])
def task():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            flash('Excel file is required!')
            return redirect(request.url)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        return redirect(url_for('process_task', file_path=filepath))
    return render_template('task.html')

@app.route('/process_task')
def process_task():
    file_path = request.args.get('file_path')
    if not file_path:
        return 'Missing file path', 400
    try:
        df = pd.read_excel(file_path)
        required_columns = ['Keywords', 'AVG. Search', 'Competition', 'RID']
        if not all(col in df.columns for col in required_columns):
            return 'Missing required columns in Excel file', 400
        df['AVG. Search'] = pd.to_numeric(df['AVG. Search'], errors='coerce')
        filtered_df = df[(df['AVG. Search'] >= 5000) & (df['Competition'] == 'Low')]
        # Remove the word 'Market' from the end of each keyword (if present)
        filtered_df['Keywords'] = filtered_df['Keywords'].str.replace(r'\s*Market$', '', regex=True)
        # Dummy processing: just save the filtered RIDs
        final_rids = filtered_df['RID'].tolist()
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'Task_Qualifying_Keywords_RIDs.csv')
        pd.DataFrame({'RID': final_rids}).to_csv(output_file, index=False)
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f'Error processing file: {e}', 500

@app.route('/rob', methods=['GET', 'POST'])
def rob():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            flash('Excel file is required!')
            return redirect(request.url)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        return redirect(url_for('process_rob', file_path=filepath))
    return render_template('task.html')

'''@app.route('/process_rob')
def process_rob():
    file_path = request.args.get('file_path')
    if not file_path:
        return 'Missing file path', 400
    try:
        df = pd.read_excel(file_path)
        columns_to_extract = ['Report ID', 'Report Name', 'Companies covered', 'Market Size Year 2025', 'CAGR', 'Forecast Period', 'Value Projection 2032']
        if not all(column in df.columns for column in columns_to_extract):
            missing_columns = [column for column in columns_to_extract if column not in df.columns]
            return f"Columns {', '.join(missing_columns)} not found in the DataFrame.", 400
        extracted_columns = df[columns_to_extract].copy()
        extracted_columns.rename(columns={
            'Report Name': 'Market Name',
            'Companies covered': 'Key Players',
            'Market Size Year 2025': 'Market Size in 2025'
        }, inplace=True)
        extracted_columns['Forecast Period'] = '2025 to 2032'
        extracted_columns['Market Size Year'] = 'Market Size in 2025:'
        extracted_columns['Market Size'] = extracted_columns['Market Size in 2025'].astype(str) + '; Market Size in 2032: ' + extracted_columns['Value Projection 2032'].astype(str)
        def make_prompt(row):
            return (
                f"Furthermore, we have a CTA that needs to be incorporated into the generated blog. Make sure all CTAs are added properly to ensure they are fully synced with the content and, from a lead generation perspective, the placement should be optimal. CTA context: The main link redirects to our main collateral published on our website. The Sample Request URL leads to a page where users can request a sample copy of the report, and the Buy Now URL allows users to directly purchase the report by making a payment. Ensure that the CTAs are placed correctly so they direct the reader to the appropriate webpage linked. The first CTA should be placed after the Market Size and Overview data, the second CTA after the Growth Factors section, and the third CTA after the Actionable Insights. Please do not make any changes to the provided data and links such as do not add brackets, or do not make changes in formatting style, because this blog will be directly published on PR website. CTA Links: First CTA- Explore the Entire Market Report here: https://www.coherentmarketinsights.com/market-insight/"
                f"{row['Market Name'].replace(' ', '-').lower()}-{row['Report ID']} , 2nd CTA- Request for Sample Copy of the Report here : https://www.coherentmarketinsights.com/insight/request-sample/"
                f"{row['Report ID']} and 3rd CTA- Get Instant Access! Purchase Research Report and Receive a 25% Discount: https://www.coherentmarketinsights.com/insight/buy-now/"
                f"{row['Report ID']}"
            )
        extracted_columns['Prompt'] = extracted_columns.apply(make_prompt, axis=1)
        output_df = extracted_columns[['Report ID', 'Market Name', 'Key Players', 'Market Size Year', 'Market Size', 'CAGR', 'Forecast Period', 'Prompt']]
        desired_order = ['Report ID', 'Market Name', 'Forecast Period', 'Market Size Year', 'Market Size', 'CAGR', 'Key Players', 'Prompt']
        output_df = output_df[desired_order]
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'ROB.xlsx')
        output_df.to_excel(output_file, index=False)
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f'Error processing file: {e}', 500'''


@app.route('/process_rob')
def process_rob():
    file_path = request.args.get('file_path')
    if not file_path:
        return 'Missing file path', 400
    try:
        df = pd.read_excel(file_path)
        columns_to_extract = ['Report ID', 'Report Name', 'Companies covered', 'Market Size Year 2025', 'CAGR', 'Forecast Period', 'Value Projection 2032']
        if not all(column in df.columns for column in columns_to_extract):
            missing_columns = [column for column in columns_to_extract if column not in df.columns]
            return f"Columns {', '.join(missing_columns)} not found in the DataFrame.", 400

        extracted_columns = df[columns_to_extract].copy()
        extracted_columns.rename(columns={
            'Report Name': 'Market Name',
            'Companies covered': 'Key Players',
            'Market Size Year 2025': 'Market Size in 2025'
        }, inplace=True)

        # Process the extracted data
        extracted_columns['Forecast Period'] = '2025 to 2032'
        extracted_columns['Market Size Year'] = 'Market Size in 2025:'
        extracted_columns['Market Size'] = extracted_columns['Market Size in 2025'].astype(str) + '; Market Size in 2032: ' + extracted_columns['Value Projection 2032'].astype(str)

        def make_prompt(row):
            return (
                f"Furthermore, we have a CTA that needs to be incorporated into the generated blog. "
                f"Make sure all CTAs are added properly to ensure they are fully synced with the content and, "
                f"from a lead generation perspective, the placement should be optimal. CTA context: The main link "
                f"redirects to our main collateral published on our website. The Sample Request URL leads to a page "
                f"where users can request a sample copy of the report, and the Buy Now URL allows users to directly "
                f"purchase the report by making a payment. Ensure that the CTAs are placed correctly so they direct the "
                f"reader to the appropriate webpage linked. The first CTA should be placed after the Market Size and Overview data, "
                f"the second CTA after the Growth Factors section, and the third CTA after the Actionable Insights. "
                f"Please do not make any changes to the provided data and links such as do not add brackets, or do not make changes "
                f"in formatting style, because this blog will be directly published on PR website. CTA Links: "
                f"First CTA- Explore the Entire Market Report here: https://www.coherentmarketinsights.com/market-insight/"
                f"{row['Market Name'].replace(' ', '-').lower()}-{row['Report ID']} , "
                f"2nd CTA- Request for Sample Copy of the Report here : https://www.coherentmarketinsights.com/insight/request-sample/"
                f"{row['Report ID']} and 3rd CTA- Get Instant Access! Purchase Research Report and Receive a 25% Discount: "
                f"https://www.coherentmarketinsights.com/insight/buy-now/{row['Report ID']}"
            )

        extracted_columns['Prompt'] = extracted_columns.apply(make_prompt, axis=1)
        output_df = extracted_columns[['Report ID', 'Market Name', 'Key Players', 'Market Size Year', 'Market Size', 'CAGR', 'Forecast Period', 'Prompt']]
        desired_order = ['Report ID', 'Market Name', 'Forecast Period', 'Market Size Year', 'Market Size', 'CAGR', 'Key Players', 'Prompt']
        output_df = output_df[desired_order]

        # Define the output path
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'ROB.xlsx')

        # Save the Excel file
        output_df.to_excel(output_file, index=False)

        # Delete all other files in the uploads folder except ROB.xlsx
        for fname in os.listdir(app.config['UPLOAD_FOLDER']):
            fpath = os.path.join(app.config['UPLOAD_FOLDER'], fname)
            if fname != 'ROB.xlsx' and os.path.isfile(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass

        # Return the file to the user for download
        return send_file(output_file, as_attachment=True, download_name='RO.xlsx')

    except Exception as e:
        return f'Error processing file: {e}', 500

def automation(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        # Trigger Power Automate Desktop flow before the view function
        trigger_power_automate_flow("Paid PR - Files Downloader")
        return func(*args, **kwargs)
    return wrapper

def trigger_power_automate_flow(flow_name):
    """
    Triggers a Power Automate Desktop flow by launching the PAD executable and running the specified flow.
    Args:
        flow_name (str): The name of the Power Automate Desktop flow to trigger.
    """
    pad_exe_path = r"C:\Program Files (x86)\Power Automate Desktop\PAD.Console.Host.exe"
    # Use the provided flow_name argument
    if not os.path.exists(pad_exe_path):
        print("Power Automate Desktop executable not found!")
        return
    command = f'"{pad_exe_path}" -flow "{flow_name}"'
    try:
        result = subprocess.run(command, shell=True, check=True, text=True, capture_output=True)
        print(f"Flow triggered successfully. Output: {result.stdout}")
        time.sleep(5)  # Wait for the app to fully open
        flow_button_coordinates = (463, 395)  # Example coordinates, replace with the ones you captured
        print(f"Clicking at {flow_button_coordinates}")
        pyautogui.click(flow_button_coordinates)
        print("Flow triggered successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error triggering flow: {e.stderr}")

@app.route('/trigger_automation')
@automation
def trigger_automation():
    return render_template('automation.html')

if __name__ == '__main__':
    app.run(debug=True)
