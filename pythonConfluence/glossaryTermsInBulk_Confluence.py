###############################################################################################################
#
#
# This script takes a CSV file and uses Confluence's REST API Server to add term pages to the Glossary V2. 
# The CSV file should contain the following column headers (capitalization matters!): 
# Term, Definition, Category 
# where Category is one of the following (capitalization doesn't matter): 
# Enterprise Assessment, Enterprise Property Tax, Enterprise Tools, Common Rolltypes, General Terms
#
# To run this, you need: 
# 1. A PAT or API Token 
#    For PAT: (Go to profile icon in top right -> Settings -> Personal Access Tokens -> Create token)
#    For API Token: go to https://id.atlassian.com/manage-profile/security/api-tokens -> Create API token
# 2. A CSV file containing your terms (and the path to this file) 
#    (To create CSV file: Create file in Excel, then Export -> Download as CSV UTF-8)
#
# I didn't have a way to test for Cloud, but this was implemented with the switch to cloud in mind.
# So, hopefully this would work either way. The only thing it should impact is authetication credentials.
#
# Helpful hints: 
# Comment out selected lines --> Ctrl + K [release] Ctrl + C 
# Un-comment selected lines --> Ctrl + K [release] Ctrl + U
# To run: Scroll to the TO RUN block at the very bottom and update the data
#
#
###############################################################################################################


import csv
import html
import requests

space_key = "iassupport"

# Mapping categories to labels and parent page IDs
category_mapping = {
    "enterprise assessment": { "parent_title": "Enterprise Assessment" },
    "enterprise property tax": { "parent_title": "Enterprise Property Tax" },
    "enterprise tools": { "parent_title": "Enterprise Tools" },
    "common rolltypes": { "parent_title": "Common Rolltypes" },
    "general terms": { "parent_title": "General Terms" }
}

# Helper function to preserve multiline formatting
def html_format_multiline(text):
    escaped = html.escape(text)
    return escaped.replace('\n', '<br />')

# Helper function to dynamically fetch the page ids by title
def get_pageid_by_title(title, space_key, base_url, headers, cloud, auth):
    url = f"{base_url}/rest/api/content"
    params = {
        "title": title,
        "spaceKey": space_key,
        "expand": "version"
    }
    response = requests.get(url, headers=headers, params=params, **({"auth": auth} if cloud else {}))
    if response.status_code == 200:
        results = response.json().get("results")
        if results:
            return results[0]["id"]
    print(f"Could not find page ID for title '{title}' in space '{space_key}'")
    return None


def main(cloud, email, api_token, csv_file_path):
    if cloud:
        auth = (email, api_token)
        base_url = "https://tylertech.atlassian.net/wiki"
        headers = { "Content-Type": "application/json" }
    else:
        base_url = "https://confl.tylertech.com"
        auth = None
        headers = {
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json"
        }

    # Read CSV
    with open(csv_file_path, mode='r', encoding='utf-8-sig') as csvfile:
        reader = csv.DictReader(csvfile)

        # Get the term, definition, and location for each term in the CSV
        for row in reader:
            term = html.escape(row.get("Term", "").strip())
            definition = html_format_multiline(row.get("Definition", "").strip())
            category = html.escape(row.get("Category", "").strip().lower())

            if not term or not definition or not category:
                print(f"Skipping incomplete row: {row}")
                continue

            # Get label and parent page ID from category mapping
            mapping = category_mapping.get(category)
            if not mapping:
                print(f"Warning: Category '{category}' not found in mapping. Skipping term '{term}'.")
                continue

            parent_title = mapping["parent_title"]
            parent_page_id = get_pageid_by_title(parent_title, space_key, base_url, headers, cloud, auth)

            if not parent_page_id:
                print(f"Skipping term '{term}' due to missing parent page ID.")
                continue

            # Construct Page Properties and CSS Stylesheet macro content
            page_properties_body = f"""
            <ac:structured-macro ac:name="details">
              <ac:rich-text-body>
                <table>
                  <tr><th>Definition</th></tr>
                  <tr><td>{definition}</td></tr>
                </table>
              </ac:rich-text-body>
            </ac:structured-macro>
            """

            # Payload to create Confluence page
            payload = {
                "type": "page",
                "title": term,
                "ancestors": [{"id": parent_page_id}],
                "space": {"key": space_key},
                "body": {
                    "storage": {
                        "value": page_properties_body,
                        "representation": "storage"
                    }
                }
            }

            # Create the page
            create_response = requests.post(
                f"{base_url}/rest/api/content",
                headers=headers,
                json=payload,
                **({"auth": auth} if cloud else {})
            )

            if create_response.status_code in (200, 201):
                page_id = create_response.json()["id"]
                print(f"Created page: {term} (ID: {page_id})")

                # Add labels: category-specific + 'glossary-terms'
                labels = [
                    {"prefix": "global", "name": "glossary-terms"}
                ]

                label_response = requests.post(
                    f"{base_url}/rest/api/content/{page_id}/label",
                    headers=headers,
                    json=labels,
                    **({"auth": auth} if cloud else {})
                )

                if label_response.status_code in (200, 204):
                    print(f"Added labels to: {term}")
                else:
                    print(f"Failed to add labels for: {term} ({label_response.status_code})")
                    print(label_response.text)

            else:
                print(f"Failed to create page: {term}")
                print(f"Status: {create_response.status_code}")
                print(create_response.text)



################################################################################################################################################
# ----------------------- TO RUN: Update this with your info, then click Start: ----------------------------------------------------------------
# You may need to go to: Tools -> Python -> Python Environments -> Open in PowerShell
# Run: python -m pip install requests
main(
    cloud=False, 
    email=None, 
    api_token="", 
    csv_file_path=r"C:\Users\.csv"
)
# ----------------------------------------------------------------------------------------------------------------------------------------------
################################################################################################################################################
