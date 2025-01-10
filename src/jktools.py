import datetime,  re

dateformat1 = "%d.%m.%Y"
dateformat2 = "%Y"
today = datetime.date.today()

def loadtime(tstring, ignore_characters = ["?"]):
    for c in ignore_characters:
        tstring = tstring.replace(c,"")
    t = str(tstring)
    try: return datetime.datetime.strptime( t, dateformat1 )
    except ValueError: pass
    # If format1 failed, try format 2 (year only)
    try: return datetime.datetime.strptime( t, dateformat2 )
    except ValueError: pass
    # Try m.YYYY
    try:
        m,y = [int(x) for x in t.split(".")]
        return datetime.datetime(y,m,1) # If only month-year given, uses 1st day of the month for testing
    except ValueError: raise ValueError(f"Date format not recognised for '{t}'")

def my2str(x): 
    if x is None: return ""
    else: return str(x)
    
def streq(u, w): 
    "Compare strings in a case-insensitive way" 
    # Primitive implementation
    return u.lower() == w.lower()

def  locstrip(s,ignorechars=""):
    """Normalize string: remove double spaces"""
    if not s or not isinstance(s,str): return s
    s = re.sub(r"\s+",  " ",  s)
    s = re.sub(" mlk\.?$",  " maalaiskunta",  s)
    s = re.sub(" pit\.?$",  " pitäjä",  s)
    s = re.sub(" lk.?$",  " landskommun",  s)
    s = re.sub("S[\:\.\,]t ",  "St ",  s)
    s = re.sub("St\. ",  "St ",  s)
    for ch in ignorechars: s = s.replace(ch,"")
    return s

def joinstr(x, y,  sep):
    if not x: return y
    else: return sep.join((x, y))
