import requests
import pandas as pd
import os

# Set up Jira API endpoint and credentials
jira_url = "https://jira.dormakaba.net/rest/api/2/search"

email = "eugene.smith@dormakaba.com"  # Replace with your Jira email
bearer_token = "NjM2Njk3MjY0NzY1Otab90VVXDR5HMdSl+bJm+2pXaq0"  # Replace with your Jira API token

# Prompt user for JQL query
print("Enter your JQL query (e.g., project = 'VARCONFIG'):")
jql_query = input("JQL Query: ").strip()

if not jql_query:
    print("No JQL query entered. Exiting.")
    exit()

# Define headers
headers = {
    "Accept": "application/json",
    "Authorization": f"Bearer {bearer_token}"  # Adjust this if needed for Jira Server
}

# Initialize variables for pagination
issues = []
start_at = 0
max_results = 100  # Fetch 100 issues per request (maximum allowed by Jira)

while True:
    # Set the query parameters with pagination
    query = {
        "jql": jql_query,
        "fields": ["key", "summary", "reporter"],  # Fields to retrieve
        "startAt": start_at,
        "maxResults": max_results
    }

    # Make the API request to search for issues
    response = requests.get(jira_url, headers=headers, params=query)

    # Print the response status code
    print(f"Response Status Code: {response.status_code}")

    # Try to parse the response data if it’s valid JSON
    try:
        data = response.json()

        # Check for error messages in the response
        if 'errorMessages' in data:
            print("Error Messages:", data['errorMessages'])
            break  # Exit if there are errors

        # Add the current page of issues to the list
        current_issues = data.get('issues', [])
        issues.extend(current_issues)

        # Break the loop if we've retrieved all issues
        if len(current_issues) < max_results:
            break

        # Increment start_at for the next page
        start_at += max_results

    except ValueError:
        print("Response content is not valid JSON.")
        break

# Process and save results to an Excel file
if issues:
    issue_data = []

    for issue in issues:
        # Check if 'fields' exists in the issue data
        if 'fields' in issue:
            issue_key = issue['key']
            summary = issue['fields'].get('summary', 'N/A')
            requester = issue['fields'].get('reporter', {}).get('displayName', 'N/A')
        else:
            issue_key = issue['key']
            summary = 'N/A'
            requester = 'N/A'

        issue_data.append({
            'KEY': issue_key,
            'Summary': summary,
            'Requester': requester
        })

    # Create a DataFrame from the issues list
    df = pd.DataFrame(issue_data)

    # Add a summary row with the total issue count
    summary_row = pd.DataFrame([{
        'KEY': '',
        'Summary': 'Total Issues',
        'Requester': len(issues)
    }])

    # Concatenate the DataFrame with the summary row
    df = pd.concat([df, summary_row], ignore_index=True)

    # Define the path to save the Excel file in the Downloads folder
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    excel_file_path = os.path.join(downloads_path, "jira_issues.xlsx")

    # Save the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False, engine='openpyxl')

    # Print the final issue count
    print(f"\nTotal Issues Retrieved: {len(issues)}")
    print(f"Data saved to {excel_file_path}")

else:
    print("No issues found.")
