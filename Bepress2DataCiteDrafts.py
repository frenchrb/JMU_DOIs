import configparser
import sys
import subprocess
from pathlib import Path

#input is path to Bepress spreadsheet saved as "XML Spreadsheet 2003"
def main(input):
    #read config file
    config = configparser.ConfigParser()
    config.read('local_settings.ini')
    
    input = Path(input)
    
    xsl_excel_named = 'coll_transforms/Excel2NamedXML.xsl'
    xsl_coll_transform = 'coll_transforms/ExcelNamed2DataCite_dnp201019_draftDOI.xsl'
    
    # Transform Excel XML into XML with named nodes
    print('Transforming Excel XML...')
    subprocess.call(['java', '-jar', config['Saxon']['saxon_path']+'saxon9he.jar', '-s:'+str(input), '-xsl:'+xsl_excel_named, '-o:excel_named_temp.xml', '-versionmsg:off'])
    print('Transformation complete')
    
    # Transform Excel Named XML to DataCite XML (items without DOIs only)
    # Output location and filenames are specified in XSLT
    print('Transforming Excel to DataCite XML...')
    subprocess.call(['java', '-jar', config['Saxon']['saxon_path']+'saxon9he.jar', '-s:excel_named_temp.xml', '-xsl:'+xsl_coll_transform])
    print('Transformation complete')
    
if __name__ == '__main__':
    main(sys.argv[1])
