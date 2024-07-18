from datetime import datetime
import openpyxl
import requests
import pandas as pd
from dotenv import load_dotenv
import os
# from lighthouse import LighthouseCI

load_dotenv() 


API_KEY = os.getenv("PAGESPEED_API_KEY")
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    #'Authorization': f'Bearer {API_KEY}'
}



# List of URLs to check
urls = [
    'https://nextjs.org/',
    'https://dinovix.com/en',
    # Add more URLs as needed
]

def get_page_speed_data(url, category='performance', strategy='mobile'):
    """Fetches Core Web Vitals, performance, accessibility, and SEO data for a given URL."""
    try:
        endpoint = f'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url={url}&category={category}&strategy={strategy}'
        response = requests.get(endpoint, headers=HEADERS)
        return response.json()
    except:
        print(f'Failed to fetch data for {url}.')
        return None
    


# check_api_key = lambda: (requests.get('https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=https://dinovix.com')).status_code == 200

def check_pagespeed():
    """Checks the PageSpeed result for the given URLs."""
    
    print('Checking URLs: \n')

    data = []

    for url in urls:
        print(f'\nChecking CWV & perfomance for {url}...')
        result = get_page_speed_data(url)
        
        if result is None:
            continue
        
        if 'lighthouseResult' in result:
            lcp = result['lighthouseResult']['audits']['largest-contentful-paint']['displayValue']
            fcp = result.get('lighthouseResult', {}).get('audits', {}).get('first-contentful-paint', {}).get('displayValue', '')
            #fid = result['lighthouseResult']['audits']['first-input']['displayValue']

            cls = result.get('lighthouseResult', {}).get('audits', {}).get('cumulative-layout-shift', {}).get('displayValue', '')
            performance_score = f"{str(result['lighthouseResult']['categories']['performance']['score']*100)}%"
            
            if 'accessibility' in result['lighthouseResult']['categories']:
                accessibility_score = f"{result['lighthouseResult']['categories']['accessibility']['score']*100}%"
            else:
                result = get_page_speed_data(url, "accessibility")
                accessibility_score = f"{result['lighthouseResult']['categories']['accessibility']['score']*100}%"


            if 'best-practices' in result['lighthouseResult']['categories']:
                best_practices = f"{result['lighthouseResult']['categories']['best-practices']['score']*100}%"
            else:
                result = get_page_speed_data(url, "best-practices")
                best_practices = f"{result['lighthouseResult']['categories']['best-practices']['score']*100}%"
            if 'seo' in result['lighthouseResult']['categories']:
                seo_score = f"{result['lighthouseResult']['categories']['seo']['score']*100}%"
            else:
                result = get_page_speed_data(url, "seo")
                seo_score = f"{result['lighthouseResult']['categories']['seo']['score']*100}%"
            
            # get overall loading experience message
            if 'overall_category' in result['loadingExperience']: 
                overall_message = result['loadingExperience']['overall_category']
            else:
                overall_message = ""

            data.append({
                'Website URL': url,
                'LCP score': lcp,
                'FCP score': fcp,
                'CLS score': cls,
                'Performance': performance_score,
                'Accessibility': accessibility_score,
                'SEO score': seo_score,
                'Best Bractices': best_practices,
                'Overall Loading Experience': overall_message
            })
            print(f'{url} - LCP: {lcp}, FCP: {fcp}, CLS: {cls}, Performance: {performance_score}, Accessibility: {accessibility_score}, SEO: {seo_score} \n')
            emoji = ''
            if overall_message.capitalize() == 'Slow' or overall_message == None:
                emoji = '👎'
            elif overall_message.capitalize()  == 'Average':
                emoji = '⚠️' 
            elif overall_message.capitalize()  == 'Fast':
                emoji = '👍'
            print(f'CWV & perfomance for {url} checked, experience: {overall_message or None} {emoji}.\n\n')
        else:
            print(f'CWV & perfomance for {url} not found.\n\n')
            continue

    if data:
        now = datetime.now().strftime("%Y-%m-%d_%Hh%Mm.xlsx")
        filename = f'cwv_report_{now}'

        write_to_excel_file_and_format(data, filename)

        print(f'Report saved to {filename}')


def write_to_excel_file_and_format(data, filename):
    """Write data to an Excel file and format the cells."""
    df = pd.DataFrame(data)
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'
    df.to_excel(filename, index=False)
    print(f'Data written to {filename}.')

    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    for column_cells in ws.iter_cols(max_row=1):
        max_length = 0
        for cell in column_cells:
            if cell.value is not None:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
        ws.column_dimensions[openpyxl.utils.get_column_letter(column_cells[0].column)].width = max_length

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value:
                cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
                # Overall Loading Experience cells should be colored based on their value
                
                if cell.value == 'FAST':
                    cell.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='00FF00')
                elif cell.value == 'AVERAGE':
                    cell.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='FFFF00')
                elif cell.value == 'SLOW':
                    cell.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='FF0000')
        # minimum column width should be set to fit the content
    wb.save(filename)
    wb.close()



def get_local_lighthouse_data(url = urls[0] or 'https://nextjs.org/'):
    """Attemptin to fetch Lighthouse data locally for a given URL."""
        # Initialize LighthouseCI
    lc = LighthouseCI(lighthouse_path="", chrome_path="C:/Program Files (x86)/Google/Chrome/Application/chrome.exe")

    config = {
        "extends": "lighthouse:default",
        "settings": {
            "categories": ["performance", "accessibility", "best-practices", "seo"]
        }
    }

    report = lc.run(url, config=config)

    for category in report.categories():
        print(f"{category}: {report.category(category)['score']} ({report.category(category)['displayValue']})")

