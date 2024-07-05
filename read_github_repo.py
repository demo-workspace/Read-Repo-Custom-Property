import requests
import xlwings as xw 
import os

# Method to read repo names from excel and proverty valeus
def fetch_repos():
  # Opening an excel file 
  try:
    wb = xw.Book('github-repo-info.xlsx') 
    ws = wb.sheets[0]
    for col_range in range(2,100):
      if ws.range("B"+str(col_range)).value:
        custom_properties = get_custom_properties(owner, ws.range("B"+str(col_range)).value, access_token)
        if custom_properties:
          print(custom_properties)
          for properties in custom_properties:
            if properties['property_name'].lower() == "owner":
              ws.range("C"+str(col_range)).value = properties['value']
            if properties['property_name'].lower() == "vertical":
              ws.range("D"+str(col_range)).value = properties['value']
        else:
          print("No custom properties found for the repository.")
      else:
        break
    wb.save()
  except requests.exceptions.RequestException as e:
    print(f"Error reading excel file: {e}")
    return None
  return None

# Method to fetch repo properties
def get_custom_properties(owner, repo_name, access_token):
  url = f"https://api.github.com/repos/{owner}/{repo_name}/properties/values"
  headers = {
      "Authorization": f"Bearer {access_token}",
      "Accept": "application/vnd.github+json"
  }
  try:
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()
  except requests.exceptions.RequestException as e:
    print(f"Error fetching custom properties: {e}")
    return None

owner = os.environ.get("github-owner")
access_token = os.environ.get("github-access-token")

fetch_repos()

