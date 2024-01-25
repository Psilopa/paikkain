from jkerror import jkError
from jktools import loadtime,  streq
import operator
import logging
progname = 'paikkain'
log = logging.getLogger(progname)

#----- rules and applying them ---------------------------------------------
_verbose = True
datebefore = 'datebefore'
dateafter ='dateafter'
known_test_types = ("equal", "datebefore", "dateafter","notempty")

def isempty(x): 
    if x is None: return True
    elif x  == "": return True
    else: return False
    

class singlerule():
    """Return a test result quality indicator.

    Return values: 0 = No value to test in GeoData, 1 = test success. 2 = test failed"""
    
    def __init__(self, rule_column_n, colname, rule_type):
        self.col = int(rule_column_n)  # Move to Excel-type indexing here!
        self.colname=colname
        self.lowercolname=colname.lower()
        self.type = rule_type.strip().lower()
        if self.type not in known_test_types: raise jkError(f"Unknown test type {self.type}")        

    def _equaltest(self, userval,  geoval): 
        if ( isempty(geoval) and isempty(userval) ) : return  1 # success
        elif ( isempty(geoval) and not isempty(userval) ) : return  2  # fail            
        elif isempty(userval) : return  2 # Data in geodata, but not in  user input
        elif streq(geoval,  userval): return  1 
        else: return  2

#    def _notequaltest(self, userval,  geoval): 
#        if ( isempty(geoval) and isempty(userval) ) : return 2 # fail
#        elif ( isempty(geoval) and not isempty(userval) ) : return  2  # fail            
#        elif isempty(userval) : return  1 # Data in geodata, but not in  user input: notequal
#        elif streq(geoval,  userval): return  2 
#        else: return  1

    def _notemptytest(self, userval,  geoval): 
        if isempty(userval) : return 2 # fail
        else: return  1 # success

    def _isemptytest(self, userval,  geoval): # Not currently in use
        if isempty(userval) : return 1 # success
        else: return  2 # fail

    def _timetest(self, userdata,  geodata,  cmpfnc): 
        try:
            userdate = loadtime(userdata)
            geodate = loadtime(geodata)
            cmpres = cmpfnc( userdate,  geodate )
            if cmpres: return 1
            else: return 2
        except ValueError as err:
            log.info(f"Failed to convert '{userdata}' or '{geodata}' to a Date: {err}" )
            raise err
        return 0

    def match(self, userdata,  geodata):     
        """Returns 0 if no real rest was done (ie anything matches), 1 for success, 2 for failure. 
        If matching failed because of an unknown input format,  raises a ValueError. """
        geoval = (geodata or "").strip() 
        if (geoval == "*"): return 0  # No test requested
        userval = userdata.get(self.lowercolname, '')
        if not userval: userval = ""
        userval = str(userval).strip()        
        if self.type == "equal":  
            return self._equaltest(userval,  geoval)
        elif self.type == "notempty":  
            return self._notemptytest(userval,  geoval)
        elif self.type in [datebefore,  dateafter]:             
            retval = None
            if ( isempty(geoval) and isempty(userval) ) : retval =  1 # empty matches empty
            elif ( isempty(geoval) and not isempty(userval) ) : retval =  2  # empty and non-empty: fail
            elif isempty(userval) : retval =  2 # empty and non-empty: fail
            else: # non-empty and non-empty
                if self.type == datebefore: cmpfnc = operator.le # userdate before (or equal to) geodata test date
                if self.type == dateafter: cmpfnc = operator.ge # after or equal
                retval = self._timetest(userval,  geoval,  cmpfnc)
            return retval
        else: raise jkError(f"Unknown test type {self.type } for test {self.colname}")        
    
