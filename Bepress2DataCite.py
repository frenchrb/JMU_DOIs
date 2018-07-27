import sys
import subprocess

#input is path to bepress spreadsheet saved as "XML Spreadsheet 2003"
def main(input):
    saxon_path = "C:\SaxonHE9-7-0-6J/"
    xsl_excel_named = "coll_transforms/Excel2NamedXML.xsl"
    xsl_coll_transform = "coll_transforms/ExcelNamed2DataCite_dnp201019_draftDOI.xsl"

    # Transform Excel XML into XML with named nodes
    #subprocess.call(["java", "-jar", "C:\SaxonHE9-7-0-6J/saxon9he.jar", "-s:bepress_spreadsheets/dnp20180724.xml", "-xsl:coll_transforms/Excel2NamedXML.xsl", "-o:pysaxon_dnp_excel_named.xml"])
    
    subprocess.call(["java", "-jar", saxon_path+"saxon9he.jar", "-s:"+input, "-xsl:"+xsl_excel_named, "-o:excel_named_temp.xml"])
    
    # Transform Excel Named XML to DataCite XML (items without DOIs only)
    # Output location and filenames are specified in XSLT
    subprocess.call(["java", "-jar", "C:\SaxonHE9-7-0-6J/saxon9he.jar", "-s:excel_named_temp.xml", "-xsl:"+xsl_coll_transform])
    
if __name__ == "__main__":
    main(sys.argv[1])
