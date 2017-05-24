import xlrd
import unicodecsv
from dateutil.relativedelta import relativedelta, FR
from datetime import datetime
import urllib2

def xls2csv (xls_contents, csv_filename, worksheet_name):
    # Converts an Excel file to a CSV file.
    # If the excel file has multiple worksheets, only the first worksheet is converted.
    # Uses unicodecsv, so it will handle Unicode characters.
    # Uses a recent version of xlrd, so it should handle old .xls and new .xlsx equally well.

    wb = xlrd.open_workbook(file_contents = xls_contents)
    sh = wb.sheet_by_name(worksheet_name)

    fh = open(csv_filename,"wb")
    csv_out = unicodecsv.writer(fh, encoding='utf-8')

    for row_number in xrange (sh.nrows):
        csv_out.writerow(sh.row_values(row_number))

    fh.close()
    
def latest_aip_file_link ():
    date_today = datetime.today() + relativedelta(weekday=FR(-1))
    link_date_string = date_today.strftime("%d-%b-%Y")
    link = "http://www.aip.com.au/pricing/pdf/AIP_TGP_Data_%s.xls" % link_date_string
    return link

def latest_aip_file_stream ():
    link = latest_aip_file_link()
    
    request = urllib2.Request(link, headers={'Host': 'www.aip.com.au','User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:47.0) Gecko/20100101 Firefox/47.0',
                        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8','Accept-Language':'en-US,en;q=0.5','Connection':'keep-alive'})
                        
    contents = urllib2.urlopen(request).read()    
    
    return contents

def download_aip_diesel_tgp_to_csv() :
    xls2csv( latest_aip_file_stream(), "aip_diesel_tgp_data.csv", "Diesel TGP" )
    
    print "Downloaded: " + latest_aip_file_link() + " to aip_diesel_tgp_data.csv"


download_aip_diesel_tgp_to_csv()