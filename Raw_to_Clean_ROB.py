import pandas as pd

def cleaning_raw_rob():
    # Load the Excel file
    df = pd.read_excel(r"C:\Users\akshat\Desktop\code\cleaned.xlsx", engine='openpyxl')

    # List of columns to extract
    columns_to_extract = ['Report ID', 'Report Name', 'Companies covered', 'Market Size Year 2025', 'CAGR', 'Forecast Period', 'Value Projection 2032']  # Replace with actual column names

    # Check if the specified columns exist in the DataFrame
    if all(column in df.columns for column in columns_to_extract):
        # Extract the specified columns
        extracted_columns = df[columns_to_extract]

        # Rename the columns
        extracted_columns.rename(columns={
            'Report Name': 'Market Name',
            'Companies covered': 'Key Players',
            'Market Size Year 2025': 'Market Size in 2025'
        }, inplace=True)

        # Set the Forecast Period column to a constant value
        extracted_columns['Forecast Period'] = '2025 to 2032'

        # Add a new column "Market Size Year" with constant value
        extracted_columns['Market Size Year'] = 'Market Size in 2025:'

        # Create the new "Market Size" column by combining "Market Size in 2025" and "Value Projection 2032"
        extracted_columns['Market Size'] = extracted_columns['Market Size in 2025'].astype(str) + '; Market Size in 2032: ' + extracted_columns['Value Projection 2032'].astype(str)

        # Create the "Prompt" column with the provided constant text and dynamically generated links
        extracted_columns['Prompt'] = extracted_columns.apply(
            lambda row: f"Furthermore, we have a CTA that needs to be incorporated into the generated blog. Make sure all CTAs are added properly to ensure they are fully synced with the content and, from a lead generation perspective, the placement should be optimal. CTA context: The main link redirects to our main collateral published on our website. The Sample Request URL leads to a page where users can request a sample copy of the report, and the Buy Now URL allows users to directly purchase the report by making a payment. Ensure that the CTAs are placed correctly so they direct the reader to the appropriate webpage linked. The first CTA should be placed after the Market Size and Overview data, the second CTA after the Growth Factors section, and the third CTA after the Actionable Insights. Please do not make any changes to the provided data and links such as do not add brackets, or do not make changes in formatting style, because this blog will be directly published on PR website. CTA Links: First CTA- Explore the Entire Market Report here: https://www.coherentmarketinsights.com/market-insight/{row['Market Name'].replace(' ', '-').lower()}-{row['Report ID']} , 2nd CTA- Request for Sample Copy of the Report here : https://www.coherentmarketinsights.com/insight/request-sample/{row['Report ID']} and 3rd CTA- Get Instant Access! Purchase Research Report and Receive a 25% Discount: https://www.coherentmarketinsights.com/insight/buy-now/{row['Report ID']}",
            axis=1
        )

        # Prepare the final output DataFrame
        output_df = extracted_columns[['Report ID', 'Market Name', 'Key Players', 'Market Size Year', 'Market Size', 'CAGR', 'Forecast Period', 'Prompt']]

        # Desired order of columns
        desired_order = ['Report ID', 'Market Name', 'Forecast Period', 'Market Size Year', 'Market Size', 'CAGR', 'Key Players', 'Prompt']

        # Reorder columns according to the desired order
        output_df = output_df[desired_order]

        # Save the final DataFrame to an Excel file
        output_df.to_excel('ROB.xlsx', index=False)

    else:
        # Print missing columns if any
        missing_columns = [column for column in columns_to_extract if column not in df.columns]
        print(f"Columns {', '.join(missing_columns)} not found in the DataFrame.")
