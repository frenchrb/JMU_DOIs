import argparse
import configparser
import os
import re
import subprocess
import sys
from datetime import datetime
from lxml import etree
from pathlib import Path

def main(arglist):
    #print(arglist)
    parser = argparse.ArgumentParser()
    parser.add_argument('setname', help='bepress collection setname (e.g., diss201019)')
    parser.add_argument('input', help='path to Bepress spreadsheet saved as "XML Spreadsheet 2003"')
    #parser.add_argument('output', help='save directory')
    #parser.add_argument('--production', help='production DOIs', action='store_true')
    args = parser.parse_args(arglist)
    #print(args)
    
    # Read config file and parse setnames into lists by category
    config = configparser.ConfigParser(allow_no_value=True)
    config.read('local_settings.ini')
    etd_setnames = []
    for i in config.items('ETD'):
        etd_setnames.append(i[0])
    #Add additional categories here
    
    setname = arglist[0]
    input = Path(arglist[1])
    
    #Timestamp output
    date_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(date_time)
    print()
    print('------------------------------------------------------------')
    print('------------------------------------------------------------')
    
    temp_file = Path('excel_named_temp.xml')
    # Remove temp_file if it already exists
    if temp_file.is_file():
        os.remove(str(temp_file))
    
    xsl_excel_named = Path('coll_transforms/Excel2NamedXML.xsl')
    if setname in etd_setnames:
        xsl_coll_transform = Path('coll_transforms/ExcelNamed2DataCite_etd_draftDOI.xsl')
    #Add additional categories here
    else:
        xsl_coll_transform = Path('coll_transforms/ExcelNamed2DataCite_' + setname + '_draftDOI.xsl')
    
    # Transform Excel XML into XML with named nodes
    print('Transforming Excel XML...')
    subprocess.call(['java', '-jar', config['Saxon']['saxon_path']+'saxon9he.jar', '-s:'+str(input), '-xsl:'+str(xsl_excel_named), '-o:'+str(temp_file), '-versionmsg:off'])
    print('Transformation complete')
    print()
    
    # Get URL from bepress spreadsheet
    tree = etree.parse(str(temp_file))
    # url = tree.xpath('/table/row/calc_url/text()')
    # url_setname = re.sub(r'^http:\/\/commons.lib.jmu.edu\/(.*)\/\d*$', r'\g<1>', url[0])
    url_setname = tree.xpath('/table/row/issue/text()')[0]
    
    # Check that stylesheet exists and setname input matches setname in metadata before doing transformation
    if not xsl_coll_transform.is_file():
        print('Stylesheet ' + str(xsl_coll_transform) + ' does not exist')
    elif not setname == url_setname:
        print('Provided setname does not match setname in bepress spreadsheet')
    else:
        # Transform Excel Named XML to DataCite XML (items without DOIs only)
        # Output location and filenames are specified in XSLT
        print('Transforming Excel to DataCite XML...')
        subprocess.call(['java', '-jar', config['Saxon']['saxon_path']+'saxon9he.jar', '-s:'+str(temp_file), '-xsl:'+str(xsl_coll_transform)])
        print('Transformation complete')
    print('------------------------------------------------------------')
    print('------------------------------------------------------------')
    print()
    
if __name__ == '__main__':
    main(sys.argv[1:])
