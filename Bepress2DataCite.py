import argparse
import configparser
import os
import subprocess
import sys
import xlrd
from datetime import datetime
from lxml import etree
from pathlib import Path


def bepress2xml(input):
    out_filename = 'bepress_as_xml_temp.xml'
    
    # Read bepress spreadsheet
    book_in = xlrd.open_workbook(str(input))
    sheet1 = book_in.sheet_by_index(0)  # get first sheet
    sheet1_col_headers = sheet1.row_values(0)

    try:
        doi_col_index = sheet1_col_headers.index('doi')
    except ValueError:
        print('DOI field not found in bepress metadata')
        return
    
    doi_col_values = sheet1.col_values(doi_col_index)  # includes header row

    # Create list of row indices without DOI
    no_doi = []
    for (x, y) in zip(range(len(doi_col_values)), doi_col_values):
        if y == '':
            no_doi.append(x)

    # Create XML document
    doc = etree.Element('table')
    tree = etree.ElementTree(doc)

    # Iterate over spreadsheet rows with no DOI
    for row in range(1, sheet1.nrows):
        if row in no_doi:
            element = etree.SubElement(doc, 'row', number=str(row))
            
            # For each column with data in the cell, add to XML document
            for col in range(0, len(sheet1_col_headers)):
                cell_contents = sheet1.cell(row, col).value
                if cell_contents != '':
                    # Convert Excel dates
                    if 'publication_date' in sheet1_col_headers[col] or 'embargo_date' in sheet1_col_headers[col]:
                        cell_contents = datetime(*xlrd.xldate_as_tuple(cell_contents, book_in.datemode))
                    a = etree.Element(sheet1.cell(0, col).value)
                    a.text = str(cell_contents)
                    element.append(a)
    tree.write(out_filename, xml_declaration=True, encoding='utf-8', pretty_print=True)


def main(arglist):
    parser = argparse.ArgumentParser()
    parser.add_argument('setname', help='bepress collection setname (e.g., diss201019)')
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
    
    setname = args.setname
    input = Path(args.input)
    production = args.production
    
    # Timestamp output
    date_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(date_time)
    print()
    print('------------------------------------------------------------')
    print('------------------------------------------------------------')
    
    if production:
        print()
        print('***********************************************************')
        print('                      PRODUCTION DOIS                      ')
        print('***********************************************************')
        print()

    temp_file = Path('bepress_as_xml_temp.xml')
    # Remove temp_file if it already exists
    if temp_file.is_file():
        os.remove(str(temp_file))

    xsl_coll_transform_filename = 'bepress2DataCite_'
    if setname in etd_setnames:
        xsl_coll_transform_filename += 'etd'
    # Add additional categories here
    else:
        xsl_coll_transform_filename += setname
    if production:
        xsl_coll_transform_filename += '_production.xsl'
    else:
        xsl_coll_transform_filename += '_draft.xsl'
    xsl_coll_transform = Path('coll_transforms/' + xsl_coll_transform_filename)
    
    # Transform bepress spreadsheet into XML
    print('Transforming bepress spreadsheet into XML...')
    bepress2xml(input)
    print('Transformation complete')
    print()
    
    # Get setname from bepress metadata
    tree = etree.parse(str(temp_file))
    # url = tree.xpath('/table/row/calc_url/text()')
    # url_setname = re.sub(r'^http:\/\/commons.lib.jmu.edu\/(.*)\/\d*$', r'\g<1>', url[0])
    url_setname = tree.xpath('/table/row/issue/text()')[0]
    
    # Check that stylesheet exists and setname input matches setname in metadata before doing transformation
    if not xsl_coll_transform.is_file():
        print('XSL stylesheet ' + str(xsl_coll_transform) + ' does not exist')
    elif not setname == url_setname:
        print('Provided setname does not match setname in bepress spreadsheet')
    else:
        # Transform bepress XML to DataCite XML
        # Output location and filenames are specified in XSLT
        print('Transforming bepress XML to DataCite XML...')
        subprocess.call(['java', '-jar', config['Saxon']['saxon_path']+'saxon9he.jar', '-s:'+str(temp_file),
                         '-xsl:'+str(xsl_coll_transform)])
        print('Transformation complete')
        print()
        
        if production:
            out_path = input.parent / 'DataCite_metadata'
        else:
            out_path = Path(os.getcwd()) / 'DataCite_metadata_drafts'
        print('Metadata files saved in ' + str(out_path))
    print('------------------------------------------------------------')
    print('------------------------------------------------------------')
    print()
    


if __name__ == '__main__':
    main(sys.argv[1:])
