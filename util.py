from pandas._libs.tslibs.timestamps import Timestamp

months = {
 1:"Gennaio",
 2:"Febbraio",
 3:"Marzo",
 4:"Aprile",
 5:"Maggio",
 6:"Giugno",
 7:"Luglio",
 8:"Agosto",
 9:"Settembre",
 10:"Ottobre",
 11:"Novembre",
 12:"Dicembre"   
}


def extract_month(ts:Timestamp):
    m =ts.month
    return m;
def extract_year(ts:Timestamp):
    y = ts.year
    return y
def int_month_to_str(month:int):
    return months[month]
