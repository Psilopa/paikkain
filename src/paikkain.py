# Call example:     jkgeoref setup.ini inputfile.xlsx
# Current implementation depends on dictionaries keeping their order, which is true in Python3 

# Standard lib
import sys, tomllib
import atexit,  datetime,  time,  argparse
import logging
from pathlib import Path
# Non std-lib
from openpyxl.styles import PatternFill
# My own
from jkerror import jkError
import jksheet, meta
from jktest import known_test_types 
from jktools import joinstr,  my2str
from jktyping import *

starttime = time.time()
_SUPPRESS_FILE_CREATION_FOR_TESTING = False
# Default logger, will overwrite by call to createlogger()
log = logging.getLogger("paikkain default logger")

def unique_keeporder(mylist):
    return list(dict.fromkeys(mylist)) 

class valid_commands(): # Available commands for output operations
    # These should be set in code
    replace = ""
    append = ""
    fillempty = ""
    nothing = ""
    @property
    def activeops(self) -> list[str]:
        return [self.replace, self.append, self.fillempty]
    @property
    def outputops(self) -> list[str]:
        return [self.replace, self.append, self.fillempty, self.nothing]

def createlogger(fn:Path) -> logging.Logger:
    logger = logging.getLogger(meta.name)
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(fn)
    ch = logging.StreamHandler()
    formatter = logging.Formatter('%(message)s [%(levelname)s]')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    logger.addHandler(fh)
    logger.addHandler(ch)   
    return logger

class WriteRow(Exception): pass

def onexit() -> None: 
    if log: log.info("Done")
    if ( ('outdata' in dir()) and  outdata ) : outdata.close()
    if ( ('geodata' in dir()) and  geodata ) : geodata.close()
    if ( ('indata' in dir()) and  indata ) : indata.close()
    endtime = time.time()
    if log: log.info("Time spent: %-.2f s" % (endtime - starttime) )

atexit.register(onexit)

def create_output_name(infn:str|Path, addition:str|Path) -> Path:
    infn = Path(infn)
    return infn.with_suffix(f".{addition}" + infn.suffix)
#  ------------------ main sections ---------------------
def read_TOML_config(confname:Path) -> dict[str,Any]:
    """Read a configuration file in TOML. Raise a jkError on error."""    
    if not confname.is_file(): raise jkError(f"Config file '{confname.absolute()}' does not exist or if not readable.")    
    try:
        with confname.open("rb") as f: config = tomllib.load(f)
    except tomllib.TOMLDecodeError as msg: raise jkError(f"Config file '{confname.absolute()}' parsing failed: f{msg}.")
    # UGLY HARDCODED COMPONENT REPLACEMENT!
    if not "filenames" in config["knowndatafiles"]: raise ValueError("Config file has no knowndatafiles:filenames item.")
    return config    


def loadSourceData( datasourcepaths:Sequence[Path], firstline:int ) -> Sequence[jksheet.GeoData]:
    sourcedatalist = []
    for sourcepath in datasourcepaths:
        log.info(f"Loading geodata from file {sourcepath}")
        geodata = jksheet.GeoData.fromfile(sourcepath, 
                                            firstline)     
        log.debug("Parsing rules from geodata file headers")
        sourcedatalist.append( geodata )
    return sourcedatalist

def verify_first_row(firstrow) -> None: # TODO: TYPING
    for i in range(len(firstrow)):
        if not firstrow[i]: 
            raise ValueError(f"File {indata.filename}, column {i+1}: Empty column name (first row) not allowed.")

def add_columns_to_output(firstrow, outdata, sourcedatalist, pnotecolname, pnote, append_original_geodata_to_column) -> None:
    cmd = valid_commands()

def get_arguments():
    ap =argparse.ArgumentParser(description='Georeferense Excel files with geodata information')
    ap.add_argument('conffn', metavar='conffn', nargs=1, help='configuration file name')
    ap.add_argument('input_files', metavar='input_files', nargs='+', help='input data file(s)')
    return ap.parse_args()

IF_EXISTS_INSERT = 16
IF_EXISTS_SKIP = 17
IF_EXISTS_ERROR = 18
def insertcolumns(outdata:jksheet.woExcel, newcolnames:list[str], position:int, if_exists:int=IF_EXISTS_SKIP):
    """Insert column names from a list.

    Keyword if_exists defines what happens if a name already exists."""
    for colname in newcolnames[::-1]: # Insert in reverse order to keep original order
        if (colname not in outdata.colnames) or if_exists == IF_EXISTS_INSERT:
            log.info(f"Adding column {colname} to output table.")
            outdata.addcolumn(position,  [colname])
        # These modes are UNTESTED
        elif if_exists == IF_EXISTS_SKIP:
            log.debug(f"Column {colname} is already in output table.")
        elif IF_EXISTS_ERROR:
            assert False, "Tried to insert an already existing column into the output table"

#  ------------------ main script ---------------------
def main():
    # READ CONFIGURATION AND KNOWN DATA FILES
    cmd = valid_commands()
    try: 
        # SETUP
        # Read command line
        args = get_arguments()
        executedir = Path(sys.argv[0]).parent
        log = createlogger( executedir / Path(meta.name + ".log") )
        log.info(f"Starting {meta.nameversion} on {datetime.datetime.now()}")
        input_files  = [Path(x) for x in args.input_files]

        # Read parameters from config file
        confgipath = Path(args.conffn[0]) 
        log.info( f"Reading configuration file {confgipath}"  )
        c = read_TOML_config(confgipath) # CONFIG DATA 
        # TODO: insert verifier here

        inc_sheetname = c['inputfiles'].get('sheetname', None)
        # TODO: is first line 0 or 1?
        inc_first_data_line = c['inputfiles'].get('first_data_line', 2) 
#        ignorechars = c.get('ignore_in_comparison',"")
        # Regular expression replacements
        if 'replacements' in c['inputfiles']:
            regular_subs = c['inputfiles']['replacements']
        else: regular_subs = {}

        knownd_keep = c['knowndatafiles'].get('keep_original_data_marker').lower()
        datasourcepaths = [ Path(x) for x in  c['knowndatafiles'].get('filenames') ]  
        try:
            cmd.replace = c['knowndatafiles']['cmd_replace']
            cmd.append = c['knowndatafiles']['cmd_append']
            cmd.fillempty = c['knowndatafiles']['cmd_fillempty']
            cmd.nothing = c['knowndatafiles']['cmd_nothing']
        except LookupError as msg: raise jkError(f"Command names must be defined in the config file: {msg}.")

        pnote = c['outputfiles'].get('transcribernote', "")
        # if pnote_shortpaths: Use only file name in the note written to output
        pnote_shortpaths = c['outputfiles'].get("short_filenames_in_notes", False)
        if pnote_shortpaths: # Shorten paths for output cell note
            notepaths = list([str(p.name) for p in datasourcepaths])
        else: notepaths = [str(x) for x in datasourcepaths] # Full paths
        
        pnote = pnote.format(nameversion = meta.nameversion,
                             sourcepaths= ", ".join(notepaths))

        if c['outputfiles'].get('add_date_to_note'):
            pnote = pnote + " (%s)" % datetime.date.today()
        outputformat = c['outputfiles'].get('output_format')
        outputformat = outputformat.lower()
        if outputformat not in ['fast-xlsx']: #'csv' not yet supported
            raise jkError( f"Unknown output format: {outputformat.upper()}" )
            
        log.info(f"Output format: {outputformat.upper()}")

        if pnote: pnotecolname = c['outputfiles'].get('transcribernotefield')
        else: pnotecolname = False 
        itemsep = c['outputfiles'].get('data_append_connector') + " "

        original_geodata_header = c['outputfiles'].get('original_geodata_to_column_header')
        append_original_geodata_to_column = c['outputfiles'].get('append_original_geodata_to_column',None)
        skip_if_content_columnnames = c['inputfiles'].get('skip_if_nonempty',[]) # Empty list default
        skip_if_content_columnnames = [x.lower() for x in skip_if_content_columnnames]
        log.debug(f"skipping row if content if found in columns {skip_if_content_columnnames}")

        rules = []
        # Read geodata files
        source_firstdataline = c['knowndatafiles']['first_dataline']  # MUST HAVE
        sourcedatalist = loadSourceData(datasourcepaths, source_firstdataline)
        assert len(sourcedatalist) == 1, "More than one source data file not yet supported."
        for sourcedata in sourcedatalist: # THIS WORKS ONLY IF WE HAVE 1 FILE
            rules = sourcedata.parse_rules(known_test_types) # Parse row matching rules from GeoData file header rows
    # WE ARE DONE WITH SETUP
    except (FileNotFoundError,  jkError) as err: 
        log.critical(f"{err}. Exiting.")
        sys.exit()

    # PROCESS INPUT FILES 
    for infn in input_files:   
        # Read user data file
        log.info(f"\n\nProcessing file {infn}") 
        try:
            # OPEN INPUT DATA
            indata = jksheet.roExcel(infn, inc_first_data_line)            
            # SETUP FOR OUTPUT FILE
            outfn = create_output_name(infn, c['outputfiles'].get('output_append_to_filename'))        
            if outputformat == 'fast-xlsx': 
                outfn = outfn.with_suffix(".out.xlsx") 
                outdata = jksheet.woExcel(outfn,inc_first_data_line)
                outdata.fill_edited_color = "fa867e"
            else:
                raise jkError(f"Unknown output format {outputformat}. Exiting.")
#            elif outputformat == 'csv':
#                outdata = jksheet.CSVOut.fromfile(infn, inc_sheetname) # Copy existing data
#                origfn = outfn
#                outfn = outfn.with_suffix("out.csv") 
            if outfn.exists(): 
                raise jkError(f"File {outfn} exists. Will not overwrite. Exiting.")
            indata_header = indata.next_row() # First row of indata 
            # Verify that 1st line is valid
            verify_first_row(indata_header)


            # ADD COLUMNS INTO OUTPUT IF CONFIG OR KNOWN DATA CONTAIN COLS NOT IN INPUT
            insertcolumns(outdata,indata_header, 1) 
            insert_point = c['outputfiles'].get('new_column_insertion_position')
            insert_point = min( insert_point, outdata.ncols ) # Update 
            outcolnames = []
            for f in sourcedatalist:
                outcolnames = f.output_column_names(cmd.activeops) 
                insertcolumns(outdata,outcolnames, insert_point, if_exists = IF_EXISTS_SKIP) 
                # Add output columns with active output commands (ie. replace, append)
            if pnotecolname and pnote and not outdata.hascolumn(pnotecolname): # We have a comment to write
                insertcolumns(outdata,[pnotecolname], insert_point, if_exists = IF_EXISTS_SKIP) 
            if append_original_geodata_to_column and not outdata.hascolumn(append_original_geodata_to_column):
                insertcolumns(outdata,[append_original_geodata_to_column], insert_point, if_exists = IF_EXISTS_SKIP) 


            # Step through input file and process line by line
            # 1st line (header line) has already been read
            rowcount = 2
            while not indata.end(): 
                if (rowcount % 10) == 0: log.info(f"Processing row {rowcount}") 
                origdict = indata.next_row_as_dict() 
                outdict =  origdict.copy()
                edited = { k: jksheet.OP_NONE for k in outdict.keys() } # Edit status for each item on this row
                try: 
                    if rowcount < inc_first_data_line: # If this is a header line, just write
                        raise WriteRow
                    # If line has content in specified columns already, skip to WriteRow
                    for skipname in skip_if_content_columnnames: 
                        val = str(origdict.get(skipname,""))
                        if val and val.strip(): # Has some content
                            raise WriteRow 
                    matchrows = []
                    for sourcedata in sourcedatalist:
                        matchrows.extend(
                            sourcedata.find_matches( origdict, rules, normalize_dict=regular_subs )
                            )

                    nmatch = len(matchrows)
                    if nmatch == 0:  
                        raise WriteRow
                    if nmatch > 1:  
                        log.debug(f"Found multiple matches for inputrow {rowcount}: {matchrows}. Check geodata source file. Skipping row")
                        raise WriteRow
                    # OK, so we have exactly one match
                    originaldata = [] # Kept to store original data from cells that may be replaced (for later reporting in the output)
                    mrow = matchrows[0] # index of the single matching row
                    # TODO PROBLEM: which geodata file did this come from!
                    assert len(sourcedatalist) == 1, "More than one source data file not yet supported."
                    geodata = sourcedatalist[0]
                    match = geodata.get_result_dict(mrow, cmd.activeops) # match as a colname: val dictionary
                    for colname,val in match.items():
#                        print(colname,val )
                        if my2str(val).strip().lower() == knownd_keep: 
                            continue # Overrule marker in known_data
                        # If column name is not in outdata, it is not an active output field name and can be ignored
                        if colname.lower() not in outdata.lowercolnames: continue        
                        # Copy original data to a field in the output file (not copying the output cell data into itself
                        if append_original_geodata_to_column and (colname != append_original_geodata_to_column):  
                            oval = str(origdict.get(colname,""))
                            if oval: originaldata.append(oval)
                        # OK to here
                        assert geodata is not None
                        oper = geodata.get_output_action_for_column(colname, cmd.outputops) 
                        if oper not in cmd.outputops:
                            continue # Skip column with actions that are not output operations
                        elif (oper == cmd.replace) or ( oper == cmd.fillempty and not outdict[colname] ):
                            outdict[colname] = val
                            edited[colname] = jksheet.OP_REPLACED
                        elif oper == cmd.append and val: # Append non-empty values only
                            outdict[colname] = joinstr( outdict.get(colname,"" ),  val ,  itemsep ) 
                            edited[colname] = jksheet.OP_APPENDED
                    if append_original_geodata_to_column: # Append old data to designated cell
                            origstr = original_geodata_header.format(originaldata
                                                                     = itemsep.join(originaldata) ) 
                            originaldata_col = append_original_geodata_to_column.lower()
                            existing_data_in_field = str(outdict.get(originaldata_col,"" )) 
                            outdict[originaldata_col] = joinstr(existing_data_in_field,  origstr,  ". ") 
                            edited[originaldata_col] = True
                    # If available, add note by the program to clumn 'pnotecolname'
                    if pnotecolname and pnote:
                        pnote_column = pnotecolname.lower()
                        outdict[pnote_column] = joinstr(outdict.get(pnote_column,"" ),  pnote ,  itemsep) 
                        edited[pnote_column] = True
                    raise WriteRow
                except WriteRow:
                    outdata.itersetrow(outdict,  edited)
                    next(outdata) # Move to next line in outdata
                    rowcount += 1
            log.info(f"Saving output file {outfn}")
            if not _SUPPRESS_FILE_CREATION_FOR_TESTING:
                outdata.close()
        except (jkError,  FileNotFoundError,  ValueError) as msg:
            log.critical(msg)
            sys.exit() 

    # Ask for any input before closing window
    # Now handled to a OS script wrapper (.bat on Windows)


if __name__ == '__main__':
    main()
