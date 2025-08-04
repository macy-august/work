###############################################################################################################
#
#
# This script produces a CSV file containing all the glossary terms in Glossary V2.
# It will contain the term, definition, and tab it is listed under. 
#
# To run this, you need: 
# 1. A PAT or API Token 
#    For PAT: (Go to profile icon in top right -> Settings -> Personal Access Tokens -> Create token)
#    For API Token: go to https://id.atlassian.com/manage-profile/security/api-tokens -> Create API token
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


import requests
import csv
import html
import re

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

    space_key = "iassupport"

    category_mapping = {
        "enterprise assessment": "Enterprise Assessment",
        "enterprise property tax": "Enterprise Property Tax",
        "enterprise tools": "Enterprise Tools",
        "common rolltypes": "Common Rolltypes",
        "general terms": "General Terms"
    }

    rows = []

    for category_key, parent_title in category_mapping.items():
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


################################################################################################################################################
# ----------------------- TO RUN: Update this with your info, then click Start: ----------------------------------------------------------------
# You may need to go to: Tools -> Python -> Python Environments -> Open in PowerShell
# Run: python -m pip install requests
export_glossary_to_csv(
    cloud=False,
    email="your.email@tylertech.com",
    api_token="",
    csv_file_path="exported_glossary.csv"
)
# ----------------------------------------------------------------------------------------------------------------------------------------------
################################################################################################################################################
