# A wrapper around spreadsheet-like files (or other similar objects)
import openpyxl     
import jktest,  jktools
from jkerror import jkError
import csv, logging, abc
from pathlib import Path

progname = 'paikkain'
log = logging.getLogger(progname)

editedFill = openpyxl.styles.PatternFill(start_color="fa867e", end_color= "fa867e", fill_type='solid')

    
class jkXLSfilesheet(): 
    """Note that all indexing in this class is 'Excel-style', ie. first row = 1, first column = 1"""
    def __init__(self):
#        self.colnames = None # A row of column titles
        self.origcolnames = None # A row of column titles as they are in the file
        self.lowercolnames = None # A row of column titles, lowercase
        self.rulerow = None # A row of column types/rules
        self.ws = None #  working sheet 
        self.wb =None # workbook
        self.fn = None # filename
        self.ncols = 0
        self.nrows = 0
        self.valuegrid = None # In-memory copy of cell values only for speed

    def open(self,  fn,  sheetname=None):
        """Uses first sheet is name is not given"""
        self.fn = fn
        self.valuegrid = None
        self.wb = openpyxl.load_workbook( fn,  data_only =True)  
        if not sheetname: sheetname = self.wb.sheetnames[0]
        if not sheetname in self.wb.sheetnames: 
            raise jkError(f"Unknown sheet name '{sheetname}' in file {fn}")
        self.ws = self.wb[sheetname]
        self.ncols = self.ws.max_column
        #print(f"NUMBER OF COLUMNS IN {self.fn} = {self.ncols}")
        self.nrows = self.ws.max_row
        if self.nrows < 3:  raise jkError(f"File {fn} too short, needs at least two header line and one data line")
        # Look for empty columns
        emptycols = []
        for col in range(1, self.ncols+1):
            if self.column_is_empty(col): emptycols.append(col)
        for col in emptycols[::-1] :
            log.debug("Deleting empty column #%s" % col )
            self.column_delete(col)  # Delete columns starting from the end        
        # Local memory copy of values for speed: create line 0 to match row indexing with self.ws (which indexes starting from 1)
        self._valuegrid_from_data() # Copy data to a grid for read access speed!
        self._updateheadervars()
        self._indexcolumnnames()

    def _valuegrid_from_data(self):
        self.valuegrid =  [[None]*self.nrows, ] 
        for row in self.ws.values: self.valuegrid.append([x for x in row])

    def _indexcolumnnames(self):
        self.n2c = {self.lowercolnames[n]: n  for n in range(1,self.ncols)}
        return self.n2c
        
    def close(self):
        if hasattr(self, 'wb') and self.wb: self.wb.close()
        if hasattr(self, 'nwb') and self.wb: self.nwb.close()
        self.wb = None
        
    def _updateheadervars(self):
        self.origcolnames = self.get_row_values( 1 ) 
        self.origcolnames = [ x.strip() for x in self.origcolnames ] 
        for i in range(self.ncols):
            if self.origcolnames[i] == '':
                for i in range(self.ncols): print(i+1, self.origcolnames[i]) # Output column name list before exiting with an Exception
                raise jkError(f"Empty values on first row not allowed. Every column should have a name. Check column {i+1}")
        self.lowercolnames = [ x.lower() for x in self.origcolnames ]
        self.rulerow =  [ str(x) for x in self.get_row_values( 2 ) ]
        self.rulerow = [ x.strip() for x in self.rulerow ]

    # Column operations
    
    def isempty_by_colname(self, row,  colname):        
        if colname.lower() not in self.lowercolnames: return True  # No column -> no value given
        cn = self.colnumber(colname)        
        return self.isempty(row,  cn)
            
    def hascolumn(self,  colname,  casesensitive=False):
        if not casesensitive: 
             return ( colname.lower() in self.lowercolnames )
        else: raise jkError("Case sensitive column name lookup not implemented")
        
    def column_get_values(self, col): 
        return next( self.ws.iter_cols(min_col=col, max_col=col,  values_only =True)  ) # Return first column from the iter_cols iterator

    def column_is_empty(self, col): 
        for val in self.column_get_values(col):
            if val is not None: return False
        return True

    def column_delete(self, col):  
        self.ws.delete_cols(col, 1)
        if self.valuegrid: 
           for row in self.valuegrid: del row[col] 
        self.ncols -= 1
                    
    def colnumber(self, name):
        """Number in column list corresponding to column name 'name'"""
        return self.n2c[name.lower()]
#    def originalcolname(self, n): return  self.origcolnames[n-1] # origcolnames is 0-indexed!
#    def lowercolname(self, n): return  self.lowercolnames[n-1] # lowercolnames is 0-indexed!

    # Row operations
    def get_row_values(self, row): return [x  or ''  for x in self.valuegrid[row] ]
    
    def get_row_as_dict(self, rowi): 
        """Return a row converted to a dictionary with column header (line 1): cell value pairs. Excel-style indexing, ie first row = 1"""
        return { k: v for (k, v) in zip(self.lowercolnames, self.valuegrid[rowi] ) }        

    # Cell operations
    def getvalue(self, row, col): # Todo: rename to getCellValue
        return self.valuegrid[row][col]
        
    def getvalue_by_colname(self, row, colname): 
        """Returns first occurence of colname, if it repeats!"""
        return self.getvalue(row, self.colnumber(colname) )
        
    def isempty(self, row,  col): 
        if (self.getvalue(row, col) == None) or \
             (self.getvalue(row, col) == '') or \
             (self.getvalue(row, col) == []): return True
        else: return False

    # Cell operations for multivalue cells
    # TODO: convert all getvalue-type calls to return a list of values
    def getvaluelist(self, row, col): # Todo: rename to getCellValueList
        return self.valuegrid[row][col]
        
    def getvaluelist_by_colname(self, row, colname): 
        """Returns cell values as a list (for multi-value cells)

        Returns first occurence of colname, if it repeats!"""
        return self.getvalue(row, self.colnumber(colname) )        

        
class GeoData(jkXLSfilesheet):
    first_data_row = 4 # Excel indexing, starts a 1!

    def __init__(self,  *args, **kwargs):        
        super(jkXLSfilesheet, self).__init__(*args, **kwargs)        
        
    @classmethod
    def fromfile(GeoData, fn,  sheetname):
        fp = Path(fn)
        print(fp)
        if not fp.exists(): raise jkError(f"File {fp} does not exist")
        x = GeoData()        
        x.open(fn,  sheetname)
        return x

    def parse_rules(self,rulenames):
        """Find columns with rules and store them. rulenames = a list of allowed rule names."""
        rules = []
        for i in range(0,  self.ncols): #            
            if self.lowercolnames[i] is None: continue  # Skip cols with nothing in title row            
            inrule = self.rulerow[i].strip().lower()
            if inrule not in rulenames: continue
            rule = jktest.singlerule( i, self.origcolnames[i], inrule)
            rules.append(rule)                
        return tuple(rules)
        
    def find_matches(self, datadict,  rules, ignorechars=""): 
        """Recturns a list of indices to matching rows """
        matches = []         
        # Standardize value: no double spaces, no .;:
        udd = {k: jktools.locstrip(v,ignorechars) for (k, v ) in datadict.items()}
        # Test each known georef row and look for perfect matches for all tests.         
        try:
            for row in range(self.first_data_row,self.nrows+1):
                testsuccesses = 0 
                matchall = True
                for rule in rules: 
                    test_against_value = self.getvalue( row,  rule.col ) # Get geodata value to test user data against
                    testresultcode = rule.match(udd,  test_against_value)
                    testsuccesses += testresultcode
                    if ( testresultcode == 2 ): # Failure to match
                        matchall = False 
                        break   
                if matchall and (testsuccesses > 0) :  # If testsuccess == 0, all tests defaulted to success because there was no data to test                
#                    matches.append(self.get_row_as_dict(row) )
                    matches.append( row )
        except ValueError: pass
        return matches
        
    def get_output_action_for_column(self, column_name,  acceptedtypes):
        # Find in which column it occurs in with an accepted output type
        column_name = column_name.lower()
        for n in range(len(self.lowercolnames)):
            outaction = self.rulerow[n]
            if (self.lowercolnames[n] == column_name) and (outaction in acceptedtypes): 
                return outaction
        return None # If no suitable column was found
    
    def output_column_names(self,  actions=[]):
        # Look throiugh
        names = []
        for i in range(len(self.rulerow)):
            if self.rulerow[i] in actions: names.append(self.origcolnames[i])
        return names

   
    
# ------------------NEW CODE BELOW THIS ------------------

#Could be used to format cells depending on the operation type
op_replaced = 1
op_appended = 2
validops = [op_replaced, op_appended]

class jkExcel(abc.ABC):
    """First row, first col = 1"""
    def __init__(self,filename,first_data_line):     
        self.fp = Path(filename)
        self._name2col = {} 
        self._lowern2col = {} # Should use lowercase column names
        self.crow = 1 # Current row
        self.wb = self._openwb()
        self.sheet = self.wb.active 
    @property
    def colnames(self):
        return self._name2col.keys()
    @property
    def lowercolnames(self):
        return self._lowern2col.keys()
    @abc.abstractmethod
    def _openwb(self): 
        # Should return an open openpyXLSX workbook
        pass
    def close(self): 
        if self.wb: self.wb.save(self.fp)
    def hascolumn(self,colname, casesensitive=False): 
        if casesensitive: return colname in self.colnames
        else: return colname.lower() in self.lowercolnames
    def _update_name2column(self):        
        _cotitles = tuple(( x.value for x in self.sheet[1]) )
        _ind = tuple(range(1,len(_cotitles)+1))
        self._name2col = { c:i for (c,i) in zip(_cotitles,_ind) if c}
        self._lowern2col = { c.lower():i for (c,i) in zip(_cotitles,_ind) if c}
    # Support simple iteration over lines
    def __next__(self): self.crow += 1
    def next(self): self.__next__()
    def end(self): 
        if self.crow > self.sheet.max_row: return True
        else: return False

# --------------------------- INPUT FILES ---------------------------
class roExcel(jkExcel):
    supports_fill = False
    def __init__(self,filename,first_data_line):     
        super().__init__(filename,first_data_line)
        self._update_name2column()
    def _openwb(self): 
        wb = openpyxl.load_workbook(self.fp) #, read_only=True
        self.rows = wb.active.values
        return wb
    def next_row(self):
        self.next()
        return next(self.rows)
    def next_row_as_dict(self):
        row = next(self.rows)
        self.next()
        return { k:v for (k,v) in zip(self.lowercolnames,row) }
#    def _getcell(self,colno,rowno): # Get value
#        return self.sheet.cell(column=colno,row=rowno).value 
#    def _getncell(self,colname,rowno): # Get value by column name
#        colno = self.name2col[colname]
#        return self._getcell(colno,rowno)
#   def iget(self,colno): return self._getcell(colno,self.crow)
#    def inget(self,colname): return self._getncell(colname,self.crow)
    
# --------------------------- OUTPUT FILES ---------------------------
class woExcel(jkExcel):
    """First row, first col = 1"""    
    supports_fill = True
    def __init__(self,filename,first_data_line):     
        self.fill_edited= None
        super().__init__(filename,first_data_line)
    def _openwb(self): 
        # Create a new workbook
        wb = openpyxl.Workbook()
        self.sheet = wb.create_sheet()
        return wb
    # PROPERTY ATTRIBUTES
    def fill_edited_color(self, color):
        self.fill_edited = openpyxl.styles.PatternFill(start_color=color, end_color= color, fill_type='solid')
    # FILE CONTENT MODIFICATION
    def addcolumn(self,position,header_rows=["New column"]):
        """Add column. Note: First column is 1 etc."""
        if header_rows[0].lower() in self._lowern2col: 
            raise ValueError(f"Column name {header_rows[0]} does already exist")
        self.sheet.insert_cols(position) 
        for ii in range(len(header_rows)):  
            self._setcell(position,1+ii,header_rows[ii])
        self._update_name2column() 
    def _setcell(self,colno,rowno,value): # Get value
        self.sheet.cell(column=colno,row=rowno).value = value
    def _setncell(self,colname,rowno,value): # Get value by column name
        colno = self._lowern2col[colname.lower()]    
        self._setcell(colno,rowno,value)
    def iterset(self,colno,value): self._setncell(colno,self.crow,value)
    def iternset(self,colname,value): self._setncell(colname,self.crow,value)
    def itersetrow(self, dict_column_and_value,  edited):
        """Create a new row from a dctionary with column names as keys
    
    The keyword edited, of provided, must be an array of len(dict_column_and_value), with values controlling cell formatting. 0 for no formatting, 
    True values (like non-zero)  integer do currently result in cess color
        """
        cellrow = [None]*len(self.colnames)
#        print("len cellrow = ",  len(cellrow))
        for colname in dict_column_and_value:       
            pos = self._lowern2col[colname.lower()] -1
#            print(pos, colname)
            cell = openpyxl.cell.WriteOnlyCell(self.sheet, value= dict_column_and_value[colname])
            cell.number_format = '@' # TEXT
            if edited and edited[colname] and self.fill_edited: cell.fill = self.fill_edited
            cellrow[pos] = cell
        self.sheet.append(cellrow)    
    def save(self): self.wb.save(str(self.fp))
 #   def newpath(self,newfp):
#        self.fp = Path(newfp)    
#    def add_note_to_filename(self,note,sep="_"):
#       newfp = self.fp.with_name(self.fp.stem + sep + note + self.fp.suffix)        
#       self.newpath(newfp)
#    def setfill(self,colno,rowno,fill): 
#        cell = self.sheet.cell(column=colno,row=rowno)        
#        cell.fill = fill
#    def itersetfill(self,colname,fill): 
#        colno = self._lowern2col[colname.lower()]    
#        self.setfill(self.crow, colno,fill)
