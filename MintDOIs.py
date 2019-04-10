import argparse
import configparser
import re
import requests
import sys
import time
import xlrd
import xlwt
import xlutils.copy
from datetime import datetime
from lxml import etree
from pathlib import Path


def main(arglist):
    parser = argparse.ArgumentParser()
    parser.add_argument('input', help='path to bepress spreadsheet in "Excel 97-2003 Workbook (.xls)" format')
    # parser.add_argument('output', help='save directory')
    parser.add_argument('--production', help='production DOIs', action='store_true')
    args = parser.parse_args(arglist)
    
    # Read config file and parse setnames into lists by category
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('local_settings.ini')
    etd_setnames = []
    for i in config.items('ETD'):
        etd_setnames.append(i[0])
    # Add additional categories here
    
    input = Path(args.input)
    production = args.production
    
    # Timestamp output
    date_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(date_time)
    
    if production:
        print()
        print('***********************************************************')
        print('                      PRODUCTION DOIS                      ')
        print('***********************************************************')
        print()
    
    # Read Bepress spreadsheet
    print()
    print('Reading spreadsheet...')
    book_in = xlrd.open_workbook(str(input))
    sheet1 = book_in.sheet_by_index(0)  # get first sheet
    # sheet1_name = book_in.sheet_names()[0]  # name of first sheet
    sheet1_col_headers = sheet1.row_values(0)
    
    try:
        doi_col_index = sheet1_col_headers.index('doi')
    except ValueError:
        print('DOI field not found in bepress metadata')
    
    url_col_index = sheet1_col_headers.index('calc_url')
    # pub_date_col_index = sheet1_col_headers.index('publication_date')
    
    # Turn the xlrd Book into xlwt Workbook
    book_out = xlutils.copy.copy(book_in)
    out_filename = input.stem + '_with_DOIs.xls'
    # out_path = input.parent / out_filename  # save in location of input spreadsheet
    out_path = Path.cwd() / 'bepress_spreadsheets' / out_filename  # save in bepress_spreadsheets folder
    
    doi_col_values = sheet1.col_values(doi_col_index)  # includes header row
    
    # Get number of blank cells in DOIs column
    doi_count = doi_col_values[1:].count('')
    if production:
        print(doi_count, 'production DOIs will be created')
    else:
        print(doi_count, 'draft DOIs will be created')
    print()
    print('------------------------------------------------------------')
    print('------------------------------------------------------------')
    time.sleep(2)
    
    # Create list of row indices of new DOIs to be created
    new_dois = []
    # print(list(zip(range(len(doi_col_values)), doi_col_values)))
    for (x, y) in zip(range(len(doi_col_values)), doi_col_values):
        # print(x,'  ',y)
        if y == '':
            new_dois.append(x)
    # print(new_dois)
    
    if production:
        metadata_dir = Path.cwd() / 'DataCite_metadata'
    else:
        metadata_dir = Path.cwd() / 'DataCite_metadata_drafts'
    
    # Iterate over spreadsheet rows with no DOI
    for row in range(1, sheet1.nrows):
        if row in new_dois:
            url = sheet1.cell(row, url_col_index).value
            set_name = re.sub(r'https:\/\/commons\.lib\.jmu\.edu\/(.*?)\/(.*)$', r'\g<1>', url, flags=re.S)
            item_number = re.sub(r'https:\/\/commons\.lib\.jmu\.edu\/(.*?)\/(.*)$', r'\g<2>', url, flags=re.S)
            
            # Get pub_year from Excel date format
            # excel_date_number = sheet1.cell(row, pub_date_col_index).value
            # pub_year, month, day, hour, minute, second = xlrd.xldate_as_tuple(excel_date_number, book_in.datemode)
            
            print('Row', row+1)
            print(sheet1.cell(row, 0).value)  # title
            print('  URL:               ' + url)
            # print('Current DOI in sheet:',sheet1.cell(row, doi_col_index).value)  # value in doi column
            
            # Construct DOI
            if production:
                doi = '10.25885/'
            else:
                doi = '10.5072/'
            if set_name in etd_setnames:
                doi += 'etd/'
            # Add additional categories here
            doi += set_name + '/' + item_number
            print('  DOI will be:       ' + doi)
            # print()
            
            metadata_filename = ''
            if set_name in etd_setnames:
                metadata_filename += 'etd_'
            # Add additional categories here
            if production:
                metadata_filename += set_name + '_' + item_number + '.xml'
            else:
                metadata_filename += set_name + '_' + item_number + '_draft.xml'
            metadata_file = metadata_dir / metadata_filename
            
            # Verify that metadata file exists
            if metadata_file in metadata_dir.iterdir():
                # Get DOI from metadata file
                tree = etree.parse(str(metadata_file))
                metadata_doi = tree.xpath('/dcite:resource/dcite:identifier[@identifierType="DOI"]/text()',
                                          namespaces={'dcite' : 'http://datacite.org/schema/kernel-4'})[0]
                print('  DOI from metadata: ' + metadata_doi)
                
                # Verify that constructed DOI matches DOI in metadata
                if doi == metadata_doi:
                    print('DOI match verified')
                    
                    # Upload DataCite metadata
                    print()
                    print('Uploading metadata (' + metadata_filename + ') to DataCite...')
                    metadata = metadata_file.open('r', encoding='utf-8').read()
                    response = requests.post(config['DataCite API']['endpoint_md'],
                                             auth=(config['DataCite API']['username'],
                                                   config['DataCite API']['password']),
                                             data=metadata.encode('utf-8'),
                                             headers={'Content-Type':'application/xml;charset=UTF-8'})
                    if response.status_code == 201:
                        print('  ' + response.text)
                        print('Metadata upload successful')
                        
                        # Send DOI/URL to DataCite to mint DOI
                        print()
                        print('Minting DOI...')
                        doi_param = 'doi=' + doi
                        url_param = 'url=' + url
                        doi_url = doi_param + '\n' + url_param
                        # print (doi_url)
                        
                        response2 = requests.put(config['DataCite API']['endpoint_doi'] + '/' + doi,
                                                 auth=(config['DataCite API']['username'],
                                                       config['DataCite API']['password']),
                                                 data=doi_url,
                                                 headers={'Content-Type':'text/plain;charset=UTF-8'})
                        if response2.status_code == 201:
                            print('  ' + response2.text)
                            print('DOI created')
                            
                            # Write DOI to sheet and save in same place as original input spreadsheet
                            book_out.get_sheet(0).write(row, doi_col_index, 'https://doi.org/' + doi)
                            book_out.save(str(out_path))
                        else:
                            print('  ' + str(response2.status_code) + ' ' + response2.text)
                            print('DOI not created')
                        
                    else:
                        print('  ' + str(response.status_code) + ' ' + response.text)
                        print('Metadata upload failed')
                else:
                    print('DOIs do not match; check metadata')
            else:
                print('Metadata file not found')
            print()
            print('------------------------------------------------------------')
    
    print('------------------------------------------------------------')
    
    # Spreadsheet with DOIs added saved in same place as original input spreadsheet
    book_out.save(str(out_path))
    print()
    print('Spreadsheet updated with DOIs')
    print('Saved at ' + str(out_path))
    print()


if __name__ == '__main__':
    main(sys.argv[1:])
