import argparse
import base64
import configparser
import requests
import sys
import xlrd
import xlwt
from datetime import datetime
from jsonmerge import merge
from jsonmerge import Merger
from pathlib import Path
from pymarc import Record, Field


def main(arglist):
# outputs:
	# short MARC records with bib number and DOI fields
    # text file of 001s
    
    parser = argparse.ArgumentParser()
    parser.add_argument('setname', help='bepress collection setname (e.g., diss201019)')
    parser.add_argument('input', help='path to bepress spreadsheet (containing DOIs) in "Excel 97-2003 Workbook (.xls)" format')
    # parser.add_argument('output', help='save directory')
    # parser.add_argument('--production', help='production DOIs', action='store_true')
    args = parser.parse_args(arglist)
    
    # Read config file and parse setnames into lists by category
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('local_settings.ini')
    etd_setnames = []
    for i in config.items('ETD'):
        etd_setnames.append(i[0])
    # Add additional categories here
    
    setname = args.setname
    input = Path(args.input)
    
    # jsonmerge setup
    schema = {"properties":{"entries":{"mergeStrategy":"append"}}}
    merger = Merger(schema)
    
    # Timestamp output
    date_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(date_time)
    print()
    print('------------------------------------------------------------')
    print('------------------------------------------------------------')
    
    # Read Bepress spreadsheet
    # TODO check that setname matches spreadsheet?
    # print()
    # print('Reading spreadsheet...')
    book_in = xlrd.open_workbook(str(input))
    sheet1 = book_in.sheet_by_index(0)  # get first sheet
    # sheet1_name = book_in.sheet_names()[0]  # name of first sheet
    sheet1_col_headers = sheet1.row_values(0)
    
    try:
        doi_col_index = sheet1_col_headers.index('doi')
    except ValueError:
        print('DOI field not found in bepress metadata')
    url_col_index = sheet1_col_headers.index('calc_url')
    
    # Read URLs and DOIs from spreadsheet
    bepress_data = {}
    for row in range(1, sheet1.nrows):
        bepress_url = sheet1.cell(row, url_col_index).value
        bepress_doi = sheet1.cell(row, doi_col_index).value
        bepress_data[bepress_url] = bepress_doi
    print(bepress_data)
    
    # Read query criteria from file, inserting setname and starting bib number
    with open('query_setname_no_doi_bib_limiter.json', 'r') as file:
        data = file.read().replace('SETNAME', setname).replace('bxxxxxxx', 'b1000000')
    # print(data)
    
    # Authenticate to get token, using Client Credentials Grant https://techdocs.iii.com/sierraapi/Content/zReference/authClient.htm
    key_secret = config.get('Sierra API', 'key') + ':' + config.get('Sierra API', 'secret')
    key_secret_encoded = base64.b64encode(key_secret.encode('UTF-8')).decode('UTF-8')
    headers = {'Authorization': 'Basic ' + key_secret_encoded,
                'Content-Type': 'application/x-www-form-urlencoded'}
    response = requests.request('POST', 'https://catalog.lib.jmu.edu/iii/sierra-api/v5/token', headers=headers)
    j = response.json()
    token = j['access_token']
    auth = 'Bearer ' + token
    headers = {
        'Accept': 'application/json',
        'Authorization': auth}
    
    # Search Sierra for records with URL+setname and no DOI in 024 field
    limit = 2000
    response = requests.request('POST', 'https://catalog.lib.jmu.edu/iii/sierra-api/v5/bibs/query?offset=0&limit=' + str(limit), headers=headers, data=data)
    # print(response.text)
    j = response.json()
    records_returned = j['total']
    # print('Records returned:', j['total'])
    j_all = j
    
    if j['total'] == 0:
        print('No ' + setname + ' records in Sierra are missing DOIs')
    else:
        # If limit was reached, repeat until all record IDs are retrieved
        while j['total'] == limit:
            # print('--------------------------------')
            last_record_id = j['entries'][-1:][0]['link'].replace('https://catalog.lib.jmu.edu/iii/sierra-api/v5/bibs/', '')
            # print('id of last record returned:', last_record_id)
            next_record_id = str(int(last_record_id) + 1)
            # print('id of starting record for next query:', next_record_id)
            
            # Read query criteria from file, inserting setname
            with open('query_setname_no_doi_bib_limiter.json', 'r') as file:
                data = file.read().replace('SETNAME', setname).replace('bxxxxxxx', 'b' + next_record_id)
            
            response = requests.request('POST', 'https://catalog.lib.jmu.edu/iii/sierra-api/v5/bibs/query?offset=0&limit=' + str(limit), headers=headers, data=data)
            j = response.json()
            records_returned += j['total']
            print('Found ' + records_returned + ' ' + setname + ' Sierra records that are missing DOIs')
            # print(response.text)
                
            # Add new response to previous ones
            j_all = merger.merge(j_all, j)
            j_all['total'] = records_returned
        # print(j_all)
        
        # Put bib IDs in list
        bib_id_list = []
        for i in j_all['entries']:
            bib_id = i['link'].replace('https://catalog.lib.jmu.edu/iii/sierra-api/v5/bibs/', '')
            bib_id_list.append(bib_id)
        # print(bib_id_list)
        
        # Get bib varField info for all records, 500 bib IDs at a time
        fields = 'varFields'
        #querystring = {'id':'3323145', 'fields':fields}
        j_data_all = {}
        records_returned_data = 0
        chunk_size = 499
        for i in range(0, len(bib_id_list), chunk_size):
            bib_id_list_partial = bib_id_list[i:i+chunk_size]
            querystring = {'id':','.join(bib_id_list_partial), 'fields':fields, 'limit':limit}
            response = requests.request('GET', 'https://catalog.lib.jmu.edu/iii/sierra-api/v5/bibs/', headers=headers, params=querystring)
            j_data = response.json()
            records_returned_data += j_data['total']
            j_data_all = merger.merge(j_data_all, j_data)
            j_data_all['total'] = records_returned_data
        
        # Parse varField data for OCLC number and URL
        sierra_data = {}
        for i in j_data_all['entries']:
            id = i['id']
            var_fields = i['varFields']
            sierra_url = ''
            
            for v in var_fields:
                if 'marcTag' in v:
                    if '001' in v['marcTag']:
                        oclc_num = v['content']
                    if '856' in v['marcTag']:
                        for s in v['subfields']:
                            if 'u' in s['tag']:
                                if 'commons.lib.jmu.edu' in s['content']:
                                    if sierra_url:
                                        sierra_url += ';'
                                    sierra_url += s['content']
            
            # Turn bib id into bib number
            bib_reversed = id[::-1]
            total = 0
            for i, digit, in enumerate(bib_reversed):
                prod = (i+2)*int(digit)
                total += prod
            checkdigit = total%11
            if checkdigit == 10:
                checkdigit = 'x'
            bib_num = 'b' + id + str(checkdigit)
            
            # print(bib_num)
            # print('OCLC number:', oclc_num)
            # print('URL:', sierra_url)
            # print()
            sierra_data[bib_num] = (oclc_num, sierra_url)
        print(sierra_data)
            
        
        # Create short MARC records with bib number and DOI fields, and TODO create spreadsheet with OCLC numbers and DOI fields
        outmarc = open('shortrecs.mrc', 'wb')
        outbook = xlwt.Workbook()
        outsheet = outbook.add_sheet('Sheet 1')
        col_headers = ['OCLC Number', 'Bib Number', '024 7_', '856 40']
        for x, y in enumerate(col_headers, 0):
            outsheet.write(0, x, y)
        outbook.save('Changes for OCLC.xls')
        
        for i, j in enumerate(sierra_data, 1):
            print(i)
            print(j)
            
            # Get DOI from spreadsheet data
            doi = bepress_data[sierra_data[j][1]]
            print(doi)
            doi_url = 'https://doi.org/' + doi
            
            spreadsheet_024 = 'doi:' + doi + ' ‡2 doi'
            spreadsheet_856 = '‡z TEXT HERE? ‡u ' + doi_url
            
            field_907 = Field(tag = '907',
                    indicators = [' ',' '],
                    subfields = ['a', '.' + j])
            field_024 = Field(tag = '024',
                    indicators = ['7',' '],
                    subfields = [
                        'a', 'doi:' + doi,
                        '2', 'doi'])
            field_856 = Field(tag = '856',
                    indicators = ['4','0'],
                    subfields = [
                        'z', 'TEXT HERE?',
                        'u', doi_url])

            record = Record()
            record.add_ordered_field(field_907)
            record.add_ordered_field(field_024)
            record.add_ordered_field(field_856)
            outmarc.write(record.as_marc())
            
            outsheet.write(i, 0, sierra_data[j][0])
            outsheet.write(i, 1, j)
            outsheet.write(i, 2, spreadsheet_024)
            outsheet.write(i, 3, spreadsheet_856)
            outbook.save('Changes for OCLC.xls') 
        outmarc.close()
    
    
if __name__ == '__main__':
    main(sys.argv[1:])
