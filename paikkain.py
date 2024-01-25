# Current implementation depends on dictionaries keeping their order, which is true in Python3 

# TODO: SHOULD CHECK THAT ALL CELLS IN FIRST 2(3) LINES HAVE SENSIBLE VALUES:
# (UNIQUE for 1, regular vocab for 2)
# TODO: Add support for locality text field manipulation even when there are existing coordinate fields
# TODO:  add support for checking Finnish bioregion names
# TODO: ADD output format replace_if_previously_empt
class WriteRow(Exception): pass

# Call example:     jkgeoref setup.ini inputfile.xlsx
from pathlib import Path
import sys,  shutil,  atexit,  tomllib,  datetime,  time,  argparse,  os
import jksheet
from jkerror import jkError
from jktest import known_test_types 
from jktools import loadtime,  joinstr,  my2str
from openpyxl.styles import PatternFill
import logging
progname = 'paikkain'

starttime = time.time()


def createlogger(fn):
    logger = logging.getLogger(progname)
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(fn)
    ch = logging.StreamHandler()
    formatter = logging.Formatter('%(message)s [%(levelname)s]')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    logger.addHandler(fh)
    logger.addHandler(ch)   
    return logger

def onexit(): 
    log.info("Done")
    if ( ('outdata' in dir()) and  outdata ) : outdata.close()
    if ( ('geodata' in dir()) and  geodata ) : geodata.close()
    if ( ('indata' in dir()) and  indata ) : indata.close()
    endtime = time.time()
    log.info(f"Time spent: %-.2f s" % (endtime - starttime) )
#    input("Hit enter to end process: ")

def dict_has_content_in(tdict,  searchkeys):
    for dc in searchkeys:
        if tdict[dc]: return True 
        else: return False
    
def row_has_date(row):
    for n in row:
        try:  
            loadtime( n ); return True # Return true is conversion to date worked
        except (ValueError,  KeyError):  pass            
    return False    

def create_output_name(infn, addition):
    infn = Path(infn)
    return infn.with_suffix(f".{addition}" + infn.suffix)

def read_TOML_config(confname):
    """Read a configuration file in TOML. Raise a jkError on error."""    
    if not conffn.is_file(): raise jkError(f"Config file '{conffn.absolute()}' does not exist or if not readable.")    
    try:
        with conffn.open("rb") as f: config = tomllib.load(f)
    except tomllib.TOMLDecodeError as msg: raise jkError(f"Config file '{conffn.absolute()}' parsing failed: f{msg}.")
    return config    

class configholder: pass # a dummy class to store config data as static properties
#  ------------------ main script
atexit.register(onexit)

# READ CONFIGURATION AND KNOWN DATA FILES
try: 
    # Read command line
    ap =argparse.ArgumentParser(description='Georeferense Excel files with geodata information')
    ap.add_argument('conffn', metavar='conffn', nargs=1, help='configuration file name')
    ap.add_argument('input_files', metavar='input_files', nargs='+', help='input data file(s)')
    args = ap.parse_args()

    executedir = Path(sys.argv[0]).parent
    log = createlogger( executedir / Path(progname + ".log") )
    log.info(f"Starting {progname} on {datetime.datetime.now()}")
    input_files  = [Path(x) for x in args.input_files]

    # Read parameters from config file
    conffn = Path(args.conffn[0]) 
    log.info( f"Reading configuration file {conffn}"  )
    c = read_TOML_config(conffn) # CONFIG DATA 

    inc = configholder()  # Config for input data files
    outc = configholder() # Config for output data
    datac = configholder() # Config for known data 

    inc.sheetname = c['inputfiles'].get('sheetname', None)
    inc.first_data_line = c['inputfiles'].get('first_data_line ', 1)

    datac.keep = c['knowndatafiles'].get('keep_original_data_marker').lower()
    ksheetname = c['knowndatafiles'].get('sheetname',None)
    kdfn = Path(c['knowndatafiles'].get('filenames')[0])  # FOR NOW, SUPPORTS ONLY 1 FILE
    cmd_replace = c['knowndatafiles'].get('cmd_replace')
    cmd_append = c['knowndatafiles'].get('cmd_append')
    cmd_nothing = c['knowndatafiles'].get('cmd_nothing')
    outputops = [cmd_replace, cmd_append, cmd_nothing]
    activeops = [cmd_replace, cmd_append]

    pnote = c['outputfiles'].get('transcribernote', "")
    output_marker = c['outputfiles'].get('filename_add')
    if c['outputfiles'].get('add_date_to_note'):
        pnote = pnote + " (%s)" % datetime.date.today()
    outputformat = c['outputfiles'].get('output_format')
    outputformat = outputformat.lower()
    if outputformat not in ['csv', 'xlsx',  'fast-xlsx']: 
        log.critical(f"Unknown output format: {outputformat.upper()}"); sys.exit()
        
    log.info(f"Output format: {outputformat.upper()}")

    if pnote: pnotecolname = c['outputfiles'].get('transcribernotefield')
    itemsep = c['outputfiles'].get('data_append_connector') + " "
    replacefillcolor = c['outputfiles'].get('replace_fillcolor')
    appendfillcolor = c['outputfiles'].get('append_fillcolor')
    new_field_insert_point = c['outputfiles'].get('new_column_insertion_position')

    original_geodata_header = c['outputfiles'].get('original_geodata_to_column_header')
    append_original_geodata_to_column = c['outputfiles'].get('append_original_geodata_to_column',None)
    skip_if_content_columnnames = c['inputfiles'].get('skip_if_nonempty')

    # Colour objects for XLS cell background setting
    replaceFill = PatternFill(start_color=replacefillcolor, end_color= replacefillcolor, fill_type='solid')
    appendFill = PatternFill(start_color=appendfillcolor, end_color= appendfillcolor, fill_type='solid')

    outdata = None
    geodata = None

    # Read geodata file
    log.info(f"Loading geodata from file {kdfn}")
    geodata = jksheet.GeoData.fromfile(Path(kdfn), ksheetname)     
    log.debug("Parsing rules from geodata file headers")
    rules = geodata.parse_rules(known_test_types) # Parse row matching rules from GeoData file header rows
    #log.debug(f"Found the following test rules:")
    for rule in rules: log.info(f"Rule for column {rule.colname}, rule type '{rule.type}'")
except (FileNotFoundError,  jkError) as err: 
    log.critical(f"{err} Exiting.")
    sys.exit()

# PROCESS INPUT FILES 
for infn in input_files:   
    # Read user data file
    log.info(f"\n\nProcessing file {infn}") 
    try:
        # Set up output file by copying in the input file
        outfn = create_output_name(infn, output_marker)        
        if outfn.exists(): 
            log.critical(f"File {outfn} exists. Will not overwrite. Exiting."); sys.exit()            
        shutil.copyfile(infn, outfn) # Make a copy of the original file, operate on it
        if outputformat == 'csv':
            outdata = jksheet.CSVOut.fromfile(outfn, inc.sheetname)
            origfn = outfn
            outfn = outfn.with_suffix("out.csv") 
            if outfn.exists(): 
                log.critical(f"File {outfn} exists. Will not overwrite. Exiting."); sys.exit()            
            outdata.outputopen(outfn)
        # THIS VERSION USE IN-PLACE EXCEL EDITING: KEEPS FORMATTING, BUT IS VERY SLOW
        elif outputformat == 'xlsx': 
            outdata = jksheet.InplaceOut.fromfile(outfn, inc.sheetname)
        elif outputformat == 'fast-xlsx': 
            outdata = jksheet.fastXLSXOut.fromfile(outfn, inc.sheetname)
            origfn = outfn
            outfn = outfn.with_suffix(".out.xlsx") 
            if outfn.exists(): 
                log.critical(f"File {outfn} exists. Will not overwrite. Exiting."); sys.exit()            
            outdata.outputopen(outfn)
            
        indata = outdata        

        # Add fields to output table as needed on the basis of geodata file header
        for colname in geodata.output_column_names(activeops):
            if not outdata.hascolumn(colname):
                log.info(f"adding column {colname} to output table")
                outdata.addcolumn( colname, new_field_insert_point )
        if pnotecolname and pnote:
            if not outdata.hascolumn(pnotecolname):
                log.info(f"adding column {pnotecolname} to output table")
                outdata.addcolumn( pnotecolname, new_field_insert_point )
        if append_original_geodata_to_column:
            if not outdata.hascolumn(append_original_geodata_to_column):
                log.info(f"adding column {append_original_geodata_to_column} to output table")
                outdata.addcolumn(append_original_geodata_to_column , new_field_insert_point )                
            original_geodata_col = outdata.colnumber(append_original_geodata_to_column) # Note; this must be last insertion, otherwise we need to update this
            
        # Copy header lines 
        outdata.copyheaders(inc.first_data_line) # Copy header lines to possible alternative output files

        # Step through input file and process line by line
        for row in range( inc.first_data_line, indata. nrows +1 ): 
            if (row % 10) == 0: log.info(f"Processing row {row}") 
            rowdict = indata.get_row_as_dict(row) 
            outdict = indata.get_row_as_dict(row) 
            edited = { k: False for k in  rowdict.keys() }
            try: 
                for skipname in skip_if_content_columnnames: 
                    if not outdata.isempty_by_colname(row,  skipname): 
                        raise WriteRow
                matchrows = geodata.find_matches(rowdict,  rules)
    #            if len(matchrows) > 1: matchrows = match_selector(geodata,  matchrows,  rowdict)
                nmatch = len(matchrows)
                if nmatch == 0:  
                    raise WriteRow
                if nmatch > 1:  
                    log.debug(f"Found multiple matches for inputrow {row}: {matchrows}. Check geodata source file. Skipping row")
                    raise WriteRow
                # OK, so we have exactly one match
                try:
                    originaldata = []
                    mrow = matchrows[0] # index of matching row
                    match = geodata.get_row_as_dict( mrow )
                    for colname,val in match.items():  # Iterate over columns in match item
                        # If column name is not in outdata, it is not an active output field name and can be ignored
                        if colname.lower() not in outdata.lowercolnames: continue                         
                        col = indata.colnumber(colname)
                        oval = indata.getvalue(row, col)  # Value in input data at this position
                        if my2str(val).strip().lower() == datac.keep: continue
                        # Copy original data to a field in the output file (not copying the output cell data into itself
                        if append_original_geodata_to_column and (col != original_geodata_col):  
                            if oval: originaldata.append( str(indata.getvalue(row, col) ) )
                        oper = geodata.get_output_action_for_column(colname, outputops) 
                        if oper not in outputops:
                            continue # Skip column with actions that are not output operations
                        elif oper == cmd_replace:
                            outdict[colname] = val
                            edited[colname] = True
                            outdata.setbackground(row, col, replaceFill)
                        elif oper == cmd_append and val: # Append non-empty values only
                            outdict[colname] = joinstr( outdict[colname] or "",  val ,  itemsep ) 
                            edited[colname] = True
                            outdata.setbackground(row, col, appendFill)
                    if append_original_geodata_to_column: # Append old data to designated cell
                            origstr = f"\n{original_geodata_header} {itemsep.join(originaldata)}" 
                            cn = append_original_geodata_to_column.lower()
                            outdict[cn] = joinstr(outdict[cn] or "",  origstr ,  "") 
                            edited[cn] = True
                            outdata.setbackground(row, original_geodata_col, appendFill)
                    # Add note by the program, if available
                    if pnotecolname and pnote:
                        cn = pnotecolname.lower()
                        outdict[cn] = joinstr(outdict[cn] or "",  pnote ,  itemsep) 
                        edited[cn] = True
                        outdata.setbackground(row, outdata.colnumber(cn), appendFill)
                    raise WriteRow
                except (jkError) as msg:
                    log.critical(msg)
                    sys.exit()
            except WriteRow: 
                outdata.writerow(row,  outdict,  edited)
        log.info(f"Saving output file {outfn}") 
        outdata.save()
        outdata.close()
        if outputformat in ['csv',  'fast-xlsx'] and origfn.exists(): os.remove(origfn) 
    except (jkError,  FileNotFoundError) as msg:
        log.critical(msg)
        sys.exit() 

# Ask for any input before closing window
# Now handled to a OS script wrapper (.bat on Windows)
