# Paikkain
Paikkain is a Excel file processor. At its core, it is a multi-cell Excel search-and-replace, where both search and replace (or append) can happen into several cells within a single row.

Data from processed files is compared line-by-line with a set of rules from a data source file: if there is exactly one match in the source, data is updated. Finally, each line (updated or not) is written into a separate output file.

It is primarily intended as an autocompletion step for Natural History collection object databasing efforts. Paikkain is run after primary data entry on semi-verbatim data which has already been split into single topic fields (like Collector or Municipality or DateBegin).

## Documentation in Finnish
See [Paikkain Wiki page](https://wiki.helsinki.fi/xwiki/bin/view/ITL/Entomology%20team%20%28Luomus%29/Hy%C3%B6nteisn%C3%A4ytteiden%20sy%C3%B6tt%C3%A4minen%20Kotka-tietokantaan%20%28Luomus%29/Paikkain/) for documentation in Finnish.

## Using paikkain
Paikkain is primarily a command line tool, but for easy of use on Windows, it comes with some BAT files that serve and drag-and-drop targets for predetermined settings.
For the Finnish Natural History Museums using Kotka as their database, it should work out of the box: just drop your Kotka data Excel on a BAT file and see what happens.

Others may need to edit the configuration files (in setup/) or the data source 

Terminology:
* Processed file - file with the data you want to enrish
* Data source file - file with already known data like coordinate data and locality identifiers for locality names.
* Configuration file - a TOML file definiing 

On the command line, the command to use is ````python src/paikkain.py -c <path_to_configuration_file>  processme1.xlsx [processme2.xlsx ...]````.

### The rules of Paikkain and the  data source file format
Paikkain looks for a single matching like the the data source file (DSF). Each row in the DSF is split into two main section: rules, and corresponding output.
* The first row is column names: these define how the rest is mapped into columns in the processed file: the column names must match exactly.
* The second tow is an action type: either a rule keyword or an output keyword like *append* or *replace*. Current supported rule keywords are *equal*, *date_before* and *date_after*.
* The third row is empty, reserved for future use
* Rules and corresponding output start from row four.
* A special keyword *<original>* in any output-type cell lets the user pass through original data even if the column type is *append* or *replace*.

Rule rule side do typically have one of four things in each cell:
* An asterisk sign (*) means any value is a valid value: the column value is not considered in comparisons.
* An empty cell means the corresponding processed data cell **must** be empty
* Text. The configurations file has some options for exactly how the match is done. By default, the comparison is case-insensitile. A list of characters to be ignored can be provided (TODO: has this been implemented already?)
* A date. This limits the time span in which the match is valid. Half-open time spans can be done by combining with two rules for a same column name.

### Updating data files.
For the latest data files relevant for insects from Finnish museum collections, see [Paikkain Google Drive](https://drive.google.com/drive/folders/19RoUMS7QUVONeXAtT8xa1tF6FYrTDYeQ?usp=sharing). For automatic updates, run ```python data_update.py``` regularly. On a command line, you can use python installer_paikkain.py --update_only.

If you want to change the update source, you have to edit the installer_paikkain.py file, which is also used as the updater.

### Modifying data files for your local setup
To modify Paikkain for your own needs, copy and edit the configuration files in /setup and the data files (by default in /paikka-aineistot). **Where content comes from, and where it goes is in general tracked by the column header values (1st row)**: Each column name (1st row) should appear only once in incoming data and in the output section of the data source file. The rules part, however, can have multiple columns with the same name. If your column naming scheme differs from that used by Kotka, update the TOML config file(s) use you and the first row of the data source files.

## Installation
* Install Python (version 3.7 or later)
* Download from https://github.com/Psilopa/paikkain (green Code button, then Download).
* Extract the ZIP package and move the paikkain-main directory to where you want it
* run the ````installer_paikkain.py```` file. Double-clicking  on it should work on Windows if you have Python installed. Or use a command prompt and run ````python installer_paikkain.py```` . It will take a while. Once you see BAT files generated, you are ready to go.
* Test by drag-dropping an excel file from the testiaineisto/ directory on top of paikkain_finland_droptarget.bat.
