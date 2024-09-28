# ScriptAccessForSqlServer

ScriptAccessForSqlServer is a simple tool to that creates a complete script of a Microsoft Access database so that it can be imported into SQL Server.  The normal SQL Server Import Access wizard will import table structure and data, but it does not import things like primary and foreign keys, defaults, indexes, etc.  This will create a complete script of the following elements:
- Table structure
- Primary Keys
- Identity fields (Autocomplete in Access)
- Default constraints
- Unique indexes
- Non-unique indexes
- Table data
- Foreign Keys

## Usage
Import the module ScriptAccessForSqlServer.bas into your Access project, and run `ScriptDatabase`.  You will be prompted for a file location, and this does the rest.