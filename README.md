# Introduction
Generate CAN dbc file with OEM defined CAN matrix (*.xls). Class `CanDatabase` represents the CAN network and the architecture is similar to Vector Candb++.

Merge CAN dbc files into a single output file. (This is where most of my energy has been applied -- soothsmith.)

# Manual
## Install
1. Put file path of 'candb.cmd' into system evironment variables.
2. Modify 'candb.py' file path in 'candb.cmd'.

## Command
Several command can be used in Command Line:
- `candb -h` show command help.
- `candb gen` generate dbc from excel.
- `candb sort` sorts a single dbc file
- `candb merge` merges multiple dbc files

### Usage
candb [-h] [-s SHEETNAME] [-t TEMPLATE] [-d] {gen} filename
- `gen` command is used to generate dbc from excel.
- `filename` the path of excle.
- `-s` specify a sheetname used in the excle workbook, optinal.
- `-t` specify a template to parse excel, optional. If not given, template is generated automatically.
- `-d` show more debug info.


candb [-h] {merge} -r filename [filename...] -o outputfilename
- `merge` command is used to merge dbc files. (does not blend with `gen`) 
- `-f` to specify a list of input files (no comma's and no repeat of the `'f`)
- `-o` to specify the name of the output file.

### Example
```C
candb gen SAIC_XXXX.xls

candb merge -f file1.dbc file2.dbc -o mergedfiles.dbc
```

## Import as module
### Use method `import_excel` to load network from excel. Parameters are defined as below:
* path:     Matrix file's path
* sheet:    Sheet name of matrix in the excel
* template: Template file which descripes matrix format<br>
### Use method `load` to load a dbc directly from a file. 
* path:     The dbc path/filename<br>
### Use method `sort` to sort by message, then signal, ascending
### Use method `save` to write to file.
* path:     The output path/filename<br>
```python
database = CanNetwork()
database.import_excel("BAIC_IPC_Matrix_CAN_20161008.xls", "IPC", "b100k_gasoline")
database.load("another_file.dbc")
database.save()
```
