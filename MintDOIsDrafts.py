import argparse
import configparser
import re
import requests
import sys
import xlrd
import xlwt
import xlutils.copy
from pathlib import Path

def main(arglist):
    #print(arglist)
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help='path to bepress spreadsheet in "Excel 97-2003 Workbook (.xls)" format')
    #parser.add_argument('output', help='save directory')
    #parser.add_argument('--production', help='production DOIs', action='store_true')
    args = parser.parse_args(arglist)
    #print(args)
    
    #Read config file and parse setnames into lists by category
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('local_settings.ini')
    etd_setnames = []
    for i in config.items('ETD'):
        etd_setnames.append(i[0])
    #Add additional categories here
    
    input = Path(arglist[0])
    
    #upload DataCite metadata containing DOI
    #sends all .xml in the specified directory
    print('Uploading metadata to DataCite...')
    metadata_dir = Path.cwd() / 'DataCite_metadata_drafts'
    for filename in metadata_dir.iterdir():
        if filename.suffix == '.xml':
            print('Sending metadata to DataCite:',filename)
            metadata = (metadata_dir / filename).open('r', encoding='utf-8').read()
            response = requests.post(config['DataCite API']['endpoint_md'], auth = (config['DataCite API']['username'],config['DataCite API']['password']), data = metadata.encode('utf-8'), headers = {'Content-Type':'application/xml;charset=UTF-8'})
            if response.status_code == 201:
                print(response.text)
            else:
                print(str(response.status_code) + ' ' + response.text)
                
    print('Metadata upload complete')
    print()
    
    #read Bepress spreadsheet
    book_in = xlrd.open_workbook(str(input))
    sheet1 = book_in.sheet_by_index(0) #get first sheet
    sheet1_name = book_in.sheet_names()[0] #name of first sheet
    #print('sheet1 type:',type(sheet1))
    sheet1_col_headers = sheet1.row_values(0)
    #print(sheet1_col_headers)
    
    try:
        doi_col_index = sheet1_col_headers.index('doi')
    except ValueError:
        print('DOI field not found in Bepress metadata')

    url_col_index = sheet1_col_headers.index('calc_url')
    pub_date_col_index = sheet1_col_headers.index('publication_date')
    
    #turn the xlrd Book into xlwt Workbook
    book_out = xlutils.copy.copy(book_in)
    
    doi_col_values = sheet1.col_values(doi_col_index) #includes header row
    #print(doi_col_values)
    
    #get number of blank cells in DOIs column
    doi_count = doi_col_values[1:].count('')
    print(doi_count,'DOIs will be created')
    print()
    
    #create list of row indices of new DOIs to be created
    new_dois = []
    #print(list(zip(range(len(doi_col_values)), doi_col_values)))
    for (x,y) in zip(range(len(doi_col_values)), doi_col_values):
        #print(x,'  ',y)
        if y == '':
            new_dois.append(x)
    #print(new_dois)
    
    #construct DOI for items without one and insert in spreadsheet
    for row in range(1,sheet1.nrows):
        if row in new_dois:
            url = sheet1.cell(row,url_col_index).value
            set_name = re.sub(r'http:\/\/commons\.lib\.jmu\.edu\/(.*?)\/(.*)$', r'\g<1>', url, flags=re.S)
            item_number = re.sub(r'http:\/\/commons\.lib\.jmu\.edu\/(.*?)\/(.*)$', r'\g<2>', url, flags=re.S)
        
            #get pub_year from Excel date format
            #excel_date_number = sheet1.cell(row,pub_date_col_index).value
            #pub_year, month, day, hour, minute, second = xlrd.xldate_as_tuple(excel_date_number, book_in.datemode)
            
            draft_doi = '10.5072/'
            if set_name in etd_setnames:
                draft_doi += 'etd/'
            #Add additional categories here
            #draft_doi = '10.5072/etd/' + set_name + '/' + item_number
            draft_doi += set_name + '/' + item_number
        
            print('Sheet row #',row+1)
            print(sheet1.cell(row, 0).value) #title
            print('URL:',url)
            #print('Current DOI in sheet:',sheet1.cell(row, doi_col_index).value) #value in doi column
            print('DOI:',draft_doi)
            #print()
            
            doi_param = 'doi='+draft_doi
            url_param = 'url='+url
            doi_uri = doi_param + '\n' + url_param
            #print (doi_uri)
    
            response2 = requests.put(config['DataCite API']['endpoint_doi'] + '/' + draft_doi, auth = (config['DataCite API']['username'], config['DataCite API']['password']), data = doi_uri, headers = {'Content-Type':'text/plain;charset=UTF-8'})
            #response2 = requests.put(config['DataCite API']['endpoint_doi'] + '/' + doi, auth = (config['DataCite API']['username'], config['DataCite API']['password']), data = doi_uri, headers = {'Content-Type':'text/plain;charset=UTF-8'})
            if response2.status_code == 201:
                print(response2.text)
                #write DOI to sheet
                book_out.get_sheet(0).write(row,doi_col_index,'https://doi.org/'+draft_doi)
            else:
                print(str(response2.status_code) + ' ' + response2.text)
                
            print()
            
    #spreadsheet with DOIs added saved in same place as original input spreadsheet
    out_filename = input.stem + '_draft_DOIs_added.xls'
    out_path = input.parent / out_filename
    book_out.save(str(out_path))
    print('Spreadsheet updated with DOIs')
    print('Saved at ' + str(out_path))
    print()
    
if __name__ == '__main__':
    main(sys.argv[1:])
