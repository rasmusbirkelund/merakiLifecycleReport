#! /usr/bin/env python3

"""Python script to generate Lifecyle Report for Meraki organizations. 
This script will create an HTML file with lists for each selected organization, which may contain Meraki hardware with EoL announcement.

Copyright (c) 2024 Rasmus Hoffmann Birkelund.
This software is licensed under ...
"""

__author__ = "Rasmus Hoffmann Birkelund <rhb@conscia.com>"
__copyright__ = "Copyright (c) Rasmus Hoffmann Birkelund."
__license__ = ""

import time
import datetime
import argparse
import pandas as pd
import meraki
import requests
import bs4 as bs
import jinja2

def GetAvailableOrganizations(p_dashboard):
    # Pick organizations you want to fetch inventory from
    print("Getting all organizations, and checking accessibility..")
    orgs = p_dashboard.organizations.getOrganizations()
    targetOrgs = []
    for org in orgs:
        if org['api']['enabled'] is True:
            if len(org['management']['details']) == 0:
                targetOrgs.append(org)
            else:
                for mgmt in org['management']['details']:
                    if mgmt['value'] == 'client allowed':
                        #print("Client allowed :)")
                        targetOrgs.append(org)
                    elif mgmt['value'] == 'client blocked':
                        print(f"\tOrg: {org['name']} - Client blocked :( -- Ignored...")
                        #print(f"\t{org['name']} access denied!\t Reason: {org['management']['details'][0]['name']}")
                    else:
                        print(f"\t{org['name']} access denied!\t Reason: {mgmt['name']} - {mgmt['value']}")
                        # raise SystemError(f"Unknown error {mgmt['name']} - {mgmt['value']}")
    print("Done!")
    return targetOrgs

def InstantiateMerakiObject(p_apikey=None):
    # Instantiate Meraki Dashboard object
    if p_apikey is None:
        print("Using with system var API Key")
        dashboard = meraki.DashboardAPI(
            suppress_logging=True
        )
    else:
        print(f"Using API Key ***{p_apikey[-4:]} ")
        dashboard = meraki.DashboardAPI(p_apikey,
            suppress_logging=True
        )
    return dashboard


def CleanUpEolTable(p_eol_df):
    # Split SKUs that had joint announcements
    pattern = 'MV21|MX64|MX65|MS220-8|series'
    mask = p_eol_df['Product'].str.contains(pattern, case=False, na=False)

    new_eol_df = p_eol_df[mask].copy()

    # Generate entries for specific submodels to count properly
    ## MX64 and MX64W
    mx64_row = new_eol_df.loc[new_eol_df['Product'].str.contains("MX64")].copy()
    mx64_row['Product'] = "MX64"
    new_eol_df.replace(to_replace='MX64, MX64W', value = 'MX64W', inplace=True)
    new_eol_df = new_eol_df._append(mx64_row)
    
    ## MV21 and MV71
    # new_eol_df.replace('MV21*', 'MV71', regex=True, inplace=True)
    mv21_row = new_eol_df.loc[new_eol_df['Product'].str.contains("MV21")].copy()
    mv21_row['Product'] = "MV21"
    new_eol_df.replace(to_replace='MV21\xa0& MV71', value = 'MV71', inplace=True)
    new_eol_df = new_eol_df._append(mv21_row)
    
    ## MS220
    new_eol_df.replace(to_replace='MS220-8', value = 'MS220-8P', inplace=True)

    ## MX65
    # new_eol_df.replace(to_replace='MX65', value = 'MX65W', inplace=True)
    mx65_row = new_eol_df.loc[new_eol_df['Product'].str.contains("MX65")].copy()
    mx65_row['Product'] = "MX65"
    new_eol_df.replace(to_replace='MX65', value = 'MX65W', inplace=True)
    new_eol_df = new_eol_df._append(mx65_row)
    
    # new_eol_df['Product'] = new_eol_df['Product'].apply(lambda x: x.strip())

    # Split up MS220 and MS320 switches in their specific submodels for proper counting
    ms220_mask = new_eol_df['Product'].str.contains('MS220\xa0series', case=False, na=False)
    ms320_mask = new_eol_df['Product'].str.contains('MS320\xa0series', case=False, na=False)
    ms220_24_row = new_eol_df[ms220_mask].copy()
    ms220_24_row["Product"]="MS220-24"
    ms220_24p_row = new_eol_df[ms220_mask].copy()
    ms220_24p_row["Product"]="MS220-24P"
    ms220_48_row = new_eol_df[ms220_mask].copy()
    ms220_48_row["Product"]="MS220-48"
    ms220_48lp_row = new_eol_df[ms220_mask].copy()
    ms220_48lp_row["Product"]="MS220-48LP"
    ms220_48fp_row = new_eol_df[ms220_mask].copy()
    ms220_48fp_row["Product"]="MS220-48FP"
    ms320_24_row = new_eol_df[ms320_mask].copy()
    ms320_24_row["Product"]="MS320-24"
    ms320_24p_row = new_eol_df[ms320_mask].copy()
    ms320_24p_row["Product"]="MS320-24P"
    ms320_48_row = new_eol_df[ms320_mask].copy()
    ms320_48_row["Product"]="MS320-48"
    ms320_48lp_row = new_eol_df[ms320_mask].copy()
    ms320_48lp_row["Product"]="MS320-48LP"
    ms320_48fp_row = new_eol_df[ms320_mask].copy()
    ms320_48fp_row["Product"]="MS320-48FP"

    # Concatenate everything
    new_eol_df = new_eol_df._append([ms220_24_row,ms220_24p_row,ms220_48_row,ms220_48lp_row,ms220_48fp_row,ms320_24_row,ms320_24p_row,ms320_48_row,ms320_48lp_row,ms320_48fp_row])
    # new_eol_df = pd.concat([new_eol_df,ms220_24_row,ms220_24p_row,ms220_48_row,ms220_48lp_row,ms220_48fp_row,ms320_24_row,ms320_24p_row,ms320_48_row,ms320_48lp_row,ms320_48fp_row])
    new_eol_df = new_eol_df[new_eol_df["Product"].str.contains("series")==False]
    # final_eol_df = pd.DataFrame()
    final_eol_df = p_eol_df._append(new_eol_df)
    # final_eol_df = pd.concat([p_eol_df, new_eol_df])
    final_eol_df.replace(to_replace='MV21\xa0& MV71', value = 'MV21', inplace=True)

    return final_eol_df

def GetEolList():
    # Get table from EoL page
    url = 'https://documentation.meraki.com/General_Administration/Other_Topics/Meraki_End-of-Life_(EOL)_Products_and_Dates'
    dfs = pd.read_html(url)

    # Get Links from table
    requested_url = requests.get(url)
    soup = bs.BeautifulSoup(requested_url.text, 'html.parser')
    table = soup.find('table')

    links = []
    for row in table.find_all('tr'):
        for td in row.find_all('td'):
            sublinks = []
            if td.find_all('a'):
                for a in td.find_all('a'):
                    sublinks.append(str(a))
                links.append(sublinks)

    # Add links to dataframe
    eol_df = dfs[0]
    eol_df['Upgrade Path'] = links
    return eol_df

def main():
    header = """Python script to generate Lifecyle Report for Meraki organizations.\n
This script will create an HTML file with lists for each selected organization, which may contain Meraki hardware with EoL announcement.
"""
    footer = f"Made by {__author__}"
    parser = argparse.ArgumentParser(
        description=header,
        epilog=footer,
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("--api_key", help="API Key")
    args = parser.parse_args()
    config = vars(args)

    dashboard = InstantiateMerakiObject(config['api_key'])

    targetOrgs = GetAvailableOrganizations(dashboard)    
    
    print("\nYour API Key has access to the following organizations:")
    i = 1
    print("0 -- All")
    for org in targetOrgs:
        print(f"{i} - {org['name']}")
        i = i+1
    choice = input("Type the number of the org or orgs (separated by commas, use - for range) that you wish to obtain lifecycle information for: ")
    if choice == '0':
        org_list = targetOrgs
    else:
        targetOrgs.insert(0,'None')
        int_choice = []
        choices = choice.split(',')
        for selection in choices:
            # selection_range = selection.split('-')
            if len(selection.split('-')) == 1: # If selection is single val, append
                int_choice.append(int(selection))
            elif len(selection.split('-')) == 2: # If seelction is range, create list with range, and append
                selection_range = [x for x in range(int(selection.split('-')[0]),int(selection.split('-')[1])+1)]
                int_choice.extend(selection_range)
        # int_choice = [int(x)-1 for x in choice.split(',')]
        org_map = map(targetOrgs.__getitem__, int_choice)
        org_list = list(org_map)

    # Filter out rows with EoL Year greater than FILTER_FOR_DATE
    FILTER_FOR_DATE = str(input("Enter year filter (press Enter for none): ")) or None

    # Get License status
    for org in org_list:
        response = dashboard.organizations.getOrganizationLicensesOverview(
            org['id']
        )
        org['licenseStatus'] = response['status']
        org['licenseExpirationDate'] = response['expirationDate']

    # Create separate lists of inventory for each org
    inventory_list = []
    for org in org_list:
        if org['licenseStatus'] in ("OK","License Required"):
            response = dashboard.organizations.getOrganizationInventoryDevices(org['id'],total_pages='all')
            if len(response) == 0:
                print(f"{org['name']} has no devices in inventory")
                inventory_list.append({
                    'id': org['id'],
                    'name': org['name'],
                    'licenseStatus': org['licenseStatus'],
                    'licenseExpirationDate': org['licenseExpirationDate'],
                    'inventory' : None
                })
            else:
                inventory_list.append({
                    'id': org['id'],
                    'name': org['name'],
                    'licenseStatus': org['licenseStatus'],
                    'licenseExpirationDate': org['licenseExpirationDate'],
                    'inventory' : response
                })
        elif org['licenseStatus'] == "License Expired":
            inventory_list.append({
                'id': org['id'],
                'name': org['name'],
                'licenseStatus': org['licenseStatus'],
                'licenseExpirationDate': org['licenseExpirationDate'],
                'inventory' : None
            })
        else:
            raise SystemError("Unknown license status")

    #inventory_list = [{f"{x['name']} - {x['id']}": dashboard.organizations.getOrganizationInventoryDevices(x['id'])} for x in org_list]

    # Get list of EOL devices
    eol_df = GetEolList()

    # Generate a new DataFrame for each Inventory List
    eol_report_list = []
    for org in inventory_list:
        if org['inventory'] is None:
            print(f"{org['name']} has no devices in inventory")
            eol_report = ""
        elif org['licenseStatus'] not in ("OK","License Required"):
            print(f"{org['name']} license not OK. Reason: {org['licenseStatus']}")
            eol_report = ""
        else:
            inventory_df = pd.DataFrame(org['inventory'])

            # Don't include devices not assigned to any networks, only consider "in use" devices
            inventory_unassigned_df = inventory_df.loc[inventory_df['networkId'].isna()].copy()
            inventory_assigned_df = inventory_df.loc[~inventory_df['networkId'].isna()].copy()

            inventory_unassigned_df['lifecycle']=""
            inventory_assigned_df['lifecycle']=""

            inventory_unassigned_df['model'].isin(eol_df['Product']).astype(int)
            inventory_assigned_df['model'].isin(eol_df['Product']).astype(int)

            final_eol_df = CleanUpEolTable(eol_df)

            final_eol_df['Unassigned Units']=final_eol_df['Product'].map(inventory_unassigned_df['model'].value_counts())
            final_eol_df['Assigned Units']=final_eol_df['Product'].map(inventory_assigned_df['model'].value_counts())
            final_eol_df = final_eol_df.fillna(0)
            final_eol_df['Total Units']=final_eol_df['Assigned Units'] + final_eol_df['Unassigned Units']
            # eol_report = final_eol_df.dropna()
            eol_report = final_eol_df[final_eol_df['Total Units'] != 0.0]
            eol_report = eol_report.sort_values(by=["Total Units"], ascending=False)

            # Filter for year, if set.
            if FILTER_FOR_DATE is not None:
                filter = eol_report['End-of-Support Date'].str[-4:] <= FILTER_FOR_DATE
                eol_report = eol_report[filter]

            # Drop index column
            eol_report = eol_report.reset_index(drop=True)

            # Construct dict of each of the reports
        eol_report_dict = {
            "name": org['name'],
            "id": org['id'],
            'licenseStatus': org['licenseStatus'],
            'licenseExpirationDate': org['licenseExpirationDate'],
            "report": eol_report
        }

        eol_report_list.append(eol_report_dict)

    # Define html document variables
    page_title_text='Cisco Meraki Lifecycle Report'
    title_text = 'Cisco Meraki Lifecycle Report'
    text = """This report lists all of your equipment currently in use that has an end of life announcement. They are ordered by the
    total units column, and the Upgrade Path column links you to the EoS announcement with recommendations on upgrade paths."""
    org_text = eol_report_list[0]['name']



    templateLoader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=templateLoader)
    TEMPLATE_FILE1 = "report_template.html"
    report = templateEnv.get_template(TEMPLATE_FILE1)
   
    # For each EOL Report, add a section to the HTML document
    entry = ""
    # for i in range(len(eol_report_list)):
    #     subheader = f'''<h2>{eol_report_list[i]['name']} -- {eol_report_list[i]['id']} -- {eol_report_list[i]['licenseExpirationDate']}</h2>'''
    #     if eol_report_list[i]['licenseStatus'] == 'OK':
    #         table = f'''{eol_report_list[i]['report'].to_html(render_links=True, escape=False, index=False)}'''
    #     elif eol_report_list[i]['licenseStatus'] == 'License Expired':
    #         table = '<h3>No information due to License Expired</h3>'
    #     entry += subheader+table
    for i in range(len(eol_report_list)):
        if isinstance(eol_report_list[i]['report'], pd.DataFrame):
            if eol_report_list[i]['report'].empty:
                entry += f'''<h2>{eol_report_list[i]['name']} -- OrgID: {eol_report_list[i]['id']} -- Expiration date: {eol_report_list[i]['licenseExpirationDate']}</h2> <p style="color:red;">Nothing to report!</p>'''
            else:
                entry += f'''<h2>{eol_report_list[i]['name']} -- OrgID: {eol_report_list[i]['id']} -- Expiration date: {eol_report_list[i]['licenseExpirationDate']}</h2> {eol_report_list[i]['report'].to_html(render_links=True, escape=False, index=False)}'''
        else:
            entry += f'''<h2>{eol_report_list[i]['name']} -- OrgID: {eol_report_list[i]['id']} -- Expiration date: {eol_report_list[i]['licenseExpirationDate']}</h2> <p style="color:red;">No information due to License Expired</p>'''
        
    settings = [
        {
            "page_title_text": page_title_text,
            "title_text": title_text, 
            "text": text,
            "org_text": org_text,
            "entry": entry
          }
    ]
        
    filename = "lifecycle_report"
    lifecycle_report = report.render(
        settings[0]
    )  # this is where to put args to the template renderer
    with open(filename+".html", mode="w", encoding="utf-8") as message:
        message.write(lifecycle_report)
        print(f"... wrote file: {filename}")
    
    # Export HTML as PDF
    # CURRENT_DIR = str(pathlib.Path().absolute())
    # filename = CURRENT_DIR+"/"+filename


    # pdfkit.from_file(filename+".html", filename+".pdf") 

    # html_to_pdf(filename, "Lifecycle Report.pdf")
    # if html_to_pdf(filename, "Lifecycle Report.pdf"):
    #     print(f"PDF generated and saved at {CURRENT_DIR}")
    # else:
    #     print("PDF generation failed")


if __name__ == "__main__":
    _script_start = time.perf_counter()
    _timestamp_ = datetime.datetime.now().strftime("%d-%m-%Y (%H:%M:%S)")
    print(f"[ ROOT ]: {_timestamp_} Script has started.")
    main()
    _script_finish = time.perf_counter()
    _timestamp_end = datetime.datetime.now().strftime("%d-%m-%Y (%H:%M:%S)")
    print(f"[ ROOT ]: {_timestamp_end} Script has finished. It took {round(_script_finish-_script_start, 2)} seconds to run.")
