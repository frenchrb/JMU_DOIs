# JMU_DOIs
App to mint DOIs from bepress Digital Commons metadata using DataCite

## Requirements
Created and tested with Python 3.5 and Saxon-HE 9.7. Requires lxml, requests, xlrd, xlwt, and xlutils. PyInstaller 3.3.1 used to create executable.


Requires a config file (local_settings.ini) in the same directory. Example of local_settings.ini:
```
[DataCite API]
endpoint_md:https://mds.datacite.org/metadata
endpoint_doi:https://mds.datacite.org/doi
username:user
password:pw

[Saxon]
saxon_path:C:\SaxonHE9-7-0-6J/

[ETD]
diss201019
dnp201019
edspec201019
master201019

[Another Category]
setname
```

## Usage
The app can be run from the command line (```python DOIminterGUI.py```) or the executable.

## Build
```pyinstaller DOIminterGUI.py --noconsole -F -n DOIminter```

## License
Apache 2.0. See [LICENSE.txt](LICENSE.txt) for more information.

## Contact
Rebecca B. French - <https://github.com/frenchrb>
