###############################################################################################################
#
#
# This script includes the functions in Confluence_BulkGlossaryTerms.sln and getAllTerms.sln
# and integrates them into a UI using ui.py that is intended to be turned into an .exe file:
# Navigate to Tools --> Python --> Python Environments --> Open in PowerShell
# Run: pip install pyinstaller
# Next, inside PowerShell, use the cd command to navigate to the folder where these files are located.
# (Example: cd C:\Users\Your.Name\OneDrive - Tyler Technologies, Inc\Desktop\Confluence\bulkTerms_Confluence\bulkTerms_Confluence)
# From here, run: pyinstaller --onefile --windowed --add-data "bulkTerms_Confluence.py;." ui.py
#
#
# Helpful hints: 
# Comment out selected lines --> Ctrl + K [release] Ctrl + C 
# Un-comment selected lines --> Ctrl + K [release] Ctrl + U
#
#
###############################################################################################################


import csv
import html
import requests
import re

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
def get_pageid_by_title(title, space_key, base_url, headers, auth, cloud):
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
            parent_page_id = get_pageid_by_title(parent_title, space_key, base_url, headers, auth, cloud)

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


# ---------- This program can verify your credentials allow you to connect to REST API ----------
def verify_rest_connection(cloud, email, token):
    if cloud:
        if not email or not token:
            raise ValueError("Email and API Token required for cloud connection.")
        auth = (email, token)
        base_url = "https://tylertech.atlassian.net/wiki"
        headers = { "Content-Type": "application/json" }
    else:
        if not token:
            raise ValueError("PAT required for server connection.")
        auth = None
        base_url = "https://confl.tylertech.com"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

    url = f"{base_url}/rest/api/user/current"
    response = requests.get(url, headers=headers, **({"auth": auth} if cloud else {}))

    if response.status_code == 200:
        try:
            user_data = response.json()
            # Check for key fields to confirm valid user data:
            if any(key in user_data for key in ["accountId", "emailAddress", "username"]):
                return True
            else:
                print("Response OK but missing user data; invalid token?")
                return False
        except Exception as e:
            print(f"Error parsing response JSON: {e}")
            return False
    else:
        print(f"Failed. Status code: {response.status_code}")
        print(response.text)
        return False


def get_child_pages(parent_page_id, base_url, headers, auth, cloud):
    url = f"{base_url}/rest/api/content/{parent_page_id}/child/page"
    response = requests.get(url, headers=headers, **({"auth": auth} if cloud else {}))
    if response.status_code == 200:
        return response.json().get("results", [])
    else:
        print(f"Failed to get child pages for parent {parent_page_id}")
        return []

def get_page_content(page_id, base_url, headers, auth, cloud):
    url = f"{base_url}/rest/api/content/{page_id}?expand=body.storage"
    response = requests.get(url, headers=headers, **({"auth": auth} if cloud else {}))
    if response.status_code == 200:
        return response.json()["body"]["storage"]["value"]
    else:
        print(f"Failed to get content for page {page_id}")
        return ""


def extract_definition_from_html(html_content):
    # Extract first <table>...</table>
    match = re.search(r"<table.*?>(.*?)</table>", html_content, re.DOTALL | re.IGNORECASE)
    if not match:
        return ""
    table_html = match.group(1)

    # Extract first <td>...</td>
    td_match = re.search(r"<td.*?>(.*?)</td>", table_html, re.DOTALL | re.IGNORECASE)
    if not td_match:
        return ""

    td_html = td_match.group(1)

    # Find all <p>...</p> inside the td
    paragraphs = re.findall(r"<p.*?>(.*?)</p>", td_html, re.DOTALL | re.IGNORECASE)

    # If no <p> found, fallback to plain text inside td
    if not paragraphs:
        # Remove any remaining tags
        text = re.sub(r"<.*?>", "", td_html)
        return html.unescape(text).strip()

    # Unescape HTML entities and strip tags inside each <p>
    clean_paragraphs = []
    for p in paragraphs:
        # Remove tags inside paragraph if any remain
        p_text = re.sub(r"<.*?>", "", p)
        p_text = html.unescape(p_text).strip()
        clean_paragraphs.append(p_text)

    # Join paragraphs with newlines
    definition = "\n".join(clean_paragraphs)

    # Normalize whitespace
    definition = re.sub(r"[ \t]+", " ", definition)
    definition = re.sub(r"\n\s*\n", "\n", definition)

    return definition.strip()


def export_glossary_to_csv(cloud, email, api_token, csv_file_path):
    if cloud:
        auth = (email, api_token)
        base_url = "https://tylertech.atlassian.net/wiki"
        headers = {"Content-Type": "application/json"}
    else:
        base_url = "https://confl.tylertech.com"
        auth = None
        headers = {
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json"
        }

    rows = []

    for category_key, mapping in category_mapping.items():
        parent_title = mapping["parent_title"]
        parent_page_id = get_pageid_by_title(parent_title, space_key, base_url, headers, auth, cloud)
        if not parent_page_id:
            print(f"Skipping category '{category_key}' due to missing parent page ID.")
            continue

        child_pages = get_child_pages(parent_page_id, base_url, headers, auth, cloud)
        print(f"Found {len(child_pages)} terms in category '{category_key}'")

        for page in child_pages:
            term = page["title"]
            page_id = page["id"]
            content_html = get_page_content(page_id, base_url, headers, auth, cloud)
            definition = extract_definition_from_html(content_html)

            rows.append({
                "Term": term,
                "Definition": definition,
                "Category": parent_title
            })

    # Write to CSV
    with open(csv_file_path, mode='w', encoding='utf-8', newline='') as csvfile:
        fieldnames = ["Term", "Definition", "Category"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    print(f"Export complete. {len(rows)} terms written to {csv_file_path}")
