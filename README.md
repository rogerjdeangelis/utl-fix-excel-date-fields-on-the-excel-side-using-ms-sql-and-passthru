# utl-fix-excel-date-fields-on-the-excel-side-using-ms-sql-and-passthru
Fix excel date fields on the excel side using ms sql and passthru   
    Fix excel date fields on the excel side using ms sql and passthru                                                    
                                                                                                                         
    If you have a primary key you could even fix the values in place with SAS passthru.                                  
                                                                                                                         
    The following code converts all the dates to character dates '01/01/1989'                                            
                                                                                                                         
    Don't bother with 'proc import' hopefully SAS will see                                                               
    the light and deprecate it before R and Python take over excel import.                                               
                                                                                                                         
    MS SQL reference on the end                                                                                          
                                                                                                                         
    github                                                                                                               
    http://tinyurl.com/y2whgvc3                                                                                          
    https://github.com/rogerjdeangelis/utl-fix-excel-date-fields-on-the-excel-side-using-ms-sql-and-passthru             
                                                                                                                         
    89 other excel repos                                                                                                 
    https://github.com/rogerjdeangelis?utf8=%E2%9C%93&tab=repositories&q=excel+in%3Aname&type=&language=                 
                                                                                                                         
    *_                   _                                                                                               
    (_)_ __  _ __  _   _| |_                                                                                             
    | | '_ \| '_ \| | | | __|                                                                                            
    | | | | | |_) | |_| | |_                                                                                             
    |_|_| |_| .__/ \__,_|\__|                                                                                            
            |_|                                                                                                          
    ;                                                                                                                    
                                                                                                                         
    Paste this into excel then apply the following formats                                                               
                                                                                                                         
    Dates       Mouse serf and apply thes formats                                                                        
                I assume thta is what exists                                                                             
                                                                                                                         
    32576       General                                                                                                  
    32607       General                                                                                                  
    32668       General                                                                                                  
    9-13-89     Text                                                                                                     
    9-15-89     Text                                                                                                     
                                                                                                                         
    *            _               _                                                                                       
      ___  _   _| |_ _ __  _   _| |_                                                                                     
     / _ \| | | | __| '_ \| | | | __|                                                                                    
    | (_) | |_| | |_| |_) | |_| | |_                                                                                     
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                    
                    |_|                                                                                                  
    ;                                                                                                                    
    DATES total obs=5                                                                                                    
                                                                                                                         
                       SAS_                                                                                              
    Obs     DATES     DATES    TEXT_DATES                                                                                
                                                                                                                         
     1     32576      10660    03/09/1989                                                                                
     2     32607      10691    04/09/1989                                                                                
     3     32668      10752    06/09/1989                                                                                
     4     9-13-89    10848    09/13/1989                                                                                
     5     9-15-89    10850    09/15/1989                                                                                
                                                                                                                         
    *          _       _   _                                                                                             
     ___  ___ | |_   _| |_(_) ___  _ __                                                                                  
    / __|/ _ \| | | | | __| |/ _ \| '_ \                                                                                 
    \__ \ (_) | | |_| | |_| | (_) | | | |                                                                                
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|                                                                                
                                                                                                                         
    ;                                                                                                                    
                                                                                                                         
    proc sql dquote=ansi;                                                                                                
       connect to excel (Path="d:\xls\dates.xlsx" mixed=yes);                                                            
         create                                                                                                          
             table dates as                                                                                              
         select                                                                                                          
            dates                                                                                                        
           ,input(dteChr,mmddyy10.)  as SAS_dates                                                                        
           ,put(calculated sas_dates,mmddyy10.) as text_dates                                                            
         from                                                                                                            
            connection to Excel                                                                                          
             (                                                                                                           
              Select                                                                                                     
                   dates                                                                                                 
                   ,iif(isnumeric(dates),format(dates,"mm/dd/yy"),cvdate(dates) ) as dteChr                              
              from                                                                                                       
                   dates                                                                                                 
             );                                                                                                          
           disconnect from Excel;                                                                                        
    quit;                                                                                                                
                                                                                                                         
                                                                                                                         
    1665  proc sql dquote=ansi;                                                                                          
    1666     connect to excel (Path="d:\xls\dates.xlsx" mixed=yes);                                                      
    NOTE: Data source is connected in READ ONLY mode.                                                                    
                                                                                                                         
    1667       create                                                                                                    
    1668           table dates as                                                                                        
    1669       select                                                                                                    
    1670          dates                                                                                                  
    1671         ,input(dteChr,mmddyy10.)  as SAS_dates                                                                  
    1672         ,put(calculated sas_dates,mmddyy10.) as text_dates                                                      
    1673       from                                                                                                      
    1674          connection to Excel                                                                                    
    1675           (                                                                                                     
    1676            Select                                                                                               
    1677                 dates                                                                                           
    1678                 ,iif(isnumeric(dates),format(dates,"mm/dd/yy"),cvdate(dates                                     
    1678! ) ) as dteChr                                                                                                  
    1679            from                                                                                                 
    1680                 dates                                                                                           
    1681           );                                                                                                    
    NOTE: Table WORK."DATES" created, with 5 rows and 3 columns.                                                         
                                                                                                                         
    1682         disconnect from Excel;                                                                                  
    1683  quit;                                                                                                          
    NOTE: PROCEDURE SQL used (Total process time):                                                                       
          real time           0.09 seconds                                                                               
          user cpu time       0.03 seconds                                                                               
          system cpu time     0.00 seconds                                                                               
          memory              3236.37k                                                                                   
          OS Memory           23020.00k                                                                                  
          Timestamp           03/15/2019 11:19:39 AM                                                                     
          Step Count                        444  Switch Count  0                                                         
                                                                                                                         
                                                                                                                         
    *                          _                                                                                         
     _ __ ___  ___   ___  __ _| |                                                                                        
    | '_ ` _ \/ __| / __|/ _` | |                                                                                        
    | | | | | \__ \ \__ \ (_| | |                                                                                        
    |_| |_| |_|___/ |___/\__, |_|                                                                                        
                            |_|                                                                                          
    ;                                                                                                                    
                                                                                                                         
                                                                                                                         
    https://ss64.com/access/                                                                                             
                                                                                                                         
    a                                                                                                                    
      Abs             The absolute value of a number (nore negative sn).                                                 
     .AddMenu         Add a custom menu bar/shortcut bar.                                                                
     .AddNew          Add a new record to a recordset.                                                                   
     .ApplyFilter     Apply a filter clause to a table, form, or report.                                                 
      Array           Create an Array.                                                                                   
      Asc             The Ascii code of a character.                                                                     
      AscW            The Unicode of a character.                                                                        
      Atn             Display the ArcTan of an angle.                                                                    
      Avg (SQL)       Average.                                                                                           
    b                                                                                                                    
     .Beep (DoCmd)    Sound a tone.                                                                                      
     .BrowseTo(DoCmd) Navate between objects.                                                                            
    c                                                                                                                    
      Call            Call a procedure.                                                                                  
     .CancelEvent (DoCmd) Cancel an event.                                                                               
     .CancelUpdate    Cancel recordset changes.                                                                          
      Case            If Then Else.                                                                                      
      CBool           Convert to boolean.                                                                                
      CByte           Convert to byte.                                                                                   
      CCur            Convert to currency (number)                                                                       
      CDate           Convert to Date.                                                                                   
      CVDate          Convert to Date.                                                                                   
      CDbl            Convert to Double (number)                                                                         
      CDec            Convert to Decimal (number)                                                                        
      Choose          Return a value from a list based on position.                                                      
      ChDir           Change the current directory or folder.                                                            
      ChDrive         Change the current drive.                                                                          
      Chr             Return a character based on an ASCII code.                                                         
     .ClearMacroError (DoCmd) Clear MacroError.                                                                          
     .Close (DoCmd)           Close a form/report/window.                                                                
     .CloseDatabase (DoCmd)   Close the database.                                                                        
      CInt                    Convert to Integer (number)                                                                
      CLng                    Convert to Long (number)                                                                   
      Command                 Return command line option string.                                                         
     .CopyDatabaseFile(DoCmd) Copy to an SQL .mdf file.                                                                  
     .CopyObject (DoCmd)      Copy an Access database object.                                                            
      Cos                     Display Cosine of an angle.                                                                
      Count (SQL)             Count records.                                                                             
      CSng             Convert to Single (number.)                                                                       
      CStr             Convert to String.                                                                                
      CurDir           Return the current path.                                                                          
      CurrentDb        Return an object variable for the current database.                                               
      CurrentUser      Return the current user.                                                                          
      CVar             Convert to a Variant.                                                                             
    d                                                                                                                    
      Date             The current date.                                                                                 
      DateAdd          Add a time interval to a date.                                                                    
      DateDiff         The time difference between two dates.                                                            
      DatePart         Return part of a given date.                                                                      
      DateSerial       Return a date given a year, month, and day.                                                       
      DateValue        Convert a string to a date.                                                                       
      DAvg             Average from a set of records.                                                                    
      Day              Return the day of the month.                                                                      
      DCount           Count the number of records in a table/query.                                                     
      Delete (SQL)          Delete records.                                                                              
     .DeleteObject (DoCmd)  Delete an object.                                                                            
      DeleteSetting         Delete a value from the users registry                                                       
     .DoMenuItem (DoCmd)    Display a menu or toolbar command.                                                           
      DFirst           The first value from a set of records.                                                            
      Dir              List the files in a folder.                                                                       
      DLast            The last value from a set of records.                                                             
      DLookup          Get the value of a particular field.                                                              
      DMax             Return the maximum value from a set of records.                                                   
      DMin             Return the minimum value from a set of records.                                                   
      DoEvents         Allow the operating system to process other events.                                               
      DStDev           Estimate Standard deviation for domain (subset of records)                                        
      DStDevP          Estimate Standard deviation for population (subset of records)                                    
      DSum             Return the sum of values from a set of records.                                                   
      DVar             Estimate variance for domain (subset of records)                                                  
      DVarP            Estimate variance for population (subset of records)                                              
    e                                                                                                                    
     .Echo             Turn screen updating on or off.                                                                   
      Environ          Return the value of an OS environment variable.                                                   
      EOF              End of file input.                                                                                
      Error            Return the error message for an error No.                                                         
      Eval             Evaluate an expression.                                                                           
      Execute(SQL/VBA) Execute a procedure or run SQL.                                                                   
      Exp              Exponential e raised to the nth power.                                                            
    f                                                                                                                    
      FileDateTime      Filename last modified date/time.                                                                
      FileLen           The size of a file in bytes.                                                                     
     .FindFirst/Last/Next/Previous Record.                                                                               
     .FindRecord(DoCmd) Find a specific record.                                                                          
      First (SQL)       Return the first value from a query.                                                             
      Fix               Return the integer portion of a number.                                                          
      For               Loop.                                                                                            
      Format            Format a Number/Date/Time.                                                                       
      FreeFile          The next file No. available to open.                                                             
      From              Specify the table(s) to be used in an .                                                          
      FV                Future Value of an annuity.                                                                      
    g                                                                                                                    
      GetAllSettings    List the settings saved in the registry.                                                         
      GetAttr           Get file/folder attributes.                                                                      
      GetObject         Return a reference to an ActiveX object                                                          
      GetSetting        Retrieve a value from the users registry.                                                        
      form.GoToPage     Move to a page on specific form.                                                                 
     .GoToRecord (DoCmd)Move to a specific record in a dataset.                                                          
    h                                                                                                                    
      Hex               Convert a number to Hex.                                                                         
      Hour              Return the hour of the day.                                                                      
     .Hourglass (DoCmd) Display the hourglass icon.                                                                      
      HyperlinkPart     Return information about data stored as a hyperlink.                                             
    i                                                                                                                    
      If Then Else      If-Then-Else                                                                                     
      IIf               If-Then-Else function.                                                                           
      Input             Return characters from a file.                                                                   
      InputBox          Prompt for user input.                                                                           
      Insert (SQL)      Add records to a table (append query).                                                           
      InStr             Return the position of one string within another.                                                
      InstrRev          Return the position of one string within another.                                                
      Int               Return the integer portion of a number.                                                          
      IPmt              Interest payment for an annuity                                                                  
      IsArray           Test if an expression is an array                                                                
      IsDate            Test if an expression is a date.                                                                 
      IsEmpty           Test if an expression is Empty (unassned).                                                       
      IsError           Test if an expression is returning an error.                                                     
      IsMissing         Test if a missing expression.                                                                    
      IsNull            Test for a NULL expression or Zero Length string.                                                
      IsNumeric         Test for a valid Number.                                                                         
      IsObject          Test if an expression is an Object.                                                              
    L                                                                                                                    
      Last (SQL)        Return the last value from a query.                                                              
      LBound            Return the smallest subscript from an array.                                                     
      LCase             Convert a string to lower-case.                                                                  
      Left              Extract a substring from a string.                                                               
      Len               Return the length of a string.                                                                   
      LoadPicture       Load a picture into an ActiveX control.                                                          
      Loc               The current position within an open file.                                                        
     .LockNavationPane(DoCmd) Lock the Navation Pane.                                                                    
      LOF               The length of a file opened with Open()                                                          
      Log               Return the natural logarithm of a number.                                                        
      LTrim             Remove leading spaces from a string.                                                             
    m                                                                                                                    
      Max (SQL)         Return the maximum value from a query.                                                           
     .Maximize (DoCmd)  Enlarge the active window.                                                                       
      Mid               Extract a substring from a string.                                                               
      Min (SQL)         Return the minimum value from a query.                                                           
     .Minimize (DoCmd)  Minimise a window.                                                                               
      Minute            Return the minute of the hour.                                                                   
      MkDir             Create directory.                                                                                
      Month             Return the month for a given date.                                                               
      MonthName         Return  a string representing the month.                                                         
     .Move              Move through a Recordset.                                                                        
     .MoveFirst/Last/Next/Previous Record                                                                                
     .MoveSize (DoCmd)  Move or Resize a Window.                                                                         
      MsgBox            Display a message in a dialogue box.                                                             
    n                                                                                                                    
      Next              Continue a for loop.                                                                             
      Now               Return the current date and time.                                                                
      Nz                Detect a NULL value or a Zero Length string.                                                     
    o                                                                                                                    
      Oct               Convert an integer to Octal.                                                                     
      OnClick, OnOpen   Events.                                                                                          
     .OpenForm (DoCmd)  Open a form.                                                                                     
     .OpenQuery (DoCmd) Open a .                                                                                         
     .OpenRecordset         Create a new Recordset.                                                                      
     .OpenReport (DoCmd)    Open a report.                                                                               
     .OutputTo (DoCmd)      Export to a Text/CSV/Spreadsheet file.                                                       
    p                                                                                                                    
      Partition (SQL)       Locate a number within a range.                                                              
     .PrintOut (DoCmd)      Print the active object (form/report etc.)                                                   
    q                                                                                                                    
      Quit                  Quit Microsoft Access                                                                        
    r                                                                                                                    
     .RefreshRecord (DoCmd) Refresh the data in a form.                                                                  
     .Rename (DoCmd)        Rename an object.                                                                            
     .RepaintObject (DoCmd) Complete any pending screen updates.                                                         
      Replace               Replace a sequence of characters in a string.                                                
     .Re               Re the data in a form or a control.                                                               
     .Restore (DoCmd)       Restore a maximized or minimized window.                                                     
      RGB                   Convert an RGB color to a number.                                                            
      Rht                 Extract a substring from a string.                                                             
      Rnd                   Generate a random number.                                                                    
      Round                 Round a number to n decimal places.                                                          
      RTrim                 Remove trailing spaces from a string.                                                        
     .RunCommand            Run an Access menu or toolbar command.                                                       
     .RunDataMacro (DoCmd)  Run a named data macro.                                                                      
     .RunMacro (DoCmd)      Run a macro.                                                                                 
     .RunSavedImportExport (DoCmd) Run a saved import or export specification.                                           
     .RunSQL (DoCmd)        Run an SQL .                                                                                 
    s                                                                                                                    
     .Save (DoCmd)          Save a database object.                                                                      
      SaveSetting           Store a value in the users registry                                                          
     .SearchForRecord(DoCmd) Search for a specific record.                                                               
      Second                Return the seconds of the minute.                                                            
      Seek                  The position within a file opened with Open.                                                 
      Select (SQL)          Retrieve data from one or more tables or queries.                                            
      Select Into (SQL)     Make-table .                                                                                 
      Select-Sub (SQL) Sub.                                                                                              
     .SelectObject (DoCmd)  Select a specific database object.                                                           
     .SendObject (DoCmd)    Send an email with a database object attached.                                               
      SendKeys              Send keystrokes to the active window.                                                        
      SetAttr               Set the attributes of a file.                                                                
     .SetDisplayedCategories (DoCmd)  Change Navation Pane display options.                                              
     .SetFilter (DoCmd)     Apply a filter to the records being displayed.                                               
      SetFocus              Move focus to a specified field or control.                                                  
     .SetMenuItem (DoCmd)   Set the state of menubar items (enabled /checked)                                            
     .SetOrderBy (DoCmd)    Apply a sort to the active datasheet, form or report.                                        
     .SetParameter (DoCmd)  Set a parameter before opening a Form or Report.                                             
     .SetWarnings (DoCmd)   Turn system messages on or off.                                                              
      Sgn                   Return the sn of a number.                                                                   
     .ShowAllRecords(DoCmd) Remove any applied filter.                                                                   
     .ShowToolbar (DoCmd)   Display or hide a custom toolbar.                                                            
      Shell                 Run an executable program.                                                                   
      Sin                   Display Sine of an angle.                                                                    
      SLN                   Straht Line Depreciation.                                                                    
      Space                 Return a number of spaces.                                                                   
      Sqr                   Return the square root of a number.                                                          
      StDev (SQL)           Estimate the standard deviation for a population.                                            
      Str                   Return a string representation of a number.                                                  
      StrComp               Compare two strings.                                                                         
      StrConv               Convert a string to Upper/lower case or Unicode.                                             
      String                Repeat a character n times.                                                                  
      Sum (SQL)             Add up the values in a  result set.                                                          
      Switch                Return one of several values.                                                                
      SysCmd                Display a progress meter.                                                                    
    t                                                                                                                    
      Top 1 *               Get first rpw                                                                                
      Tan                   Display Tangent of an angle.                                                                 
      Time                  Return the current system time.                                                              
      Timer                 Return a number (single) of seconds since midnht.                                            
      TimeSerial            Return a time given an hour, minute, and second.                                             
      TimeValue             Convert a string to a Time.                                                                  
     .TransferDatabase (DoCmd)      Import or export data to/from another database.                                      
     .TransferSharePointList(DoCmd) Import or link data from a SharePoint Foundation site.                               
     .TransferSpreadsheet (DoCmd)   Import or export data to/from a spreadsheet file.                                    
     .TransferSQLDatabase (DoCmd)   Copy an entire SQL Server database.                                                  
     .TransferText (DoCmd)          Import or export data to/from a text file.                                           
      Transform (SQL)       Create a crosstab .                                                                          
      Trim                  Remove leading and trailing spaces from a string.                                            
      TypeName              Return the data type of a variable.                                                          
    u                                                                                                                    
      UBound                Return the largest subscript from an array.                                                  
      UCase                 Convert a string to upper-case.                                                              
      Undo                  Undo the last data edit.                                                                     
      Union (SQL)           Combine the results of two SQL queries.                                                      
      Update (SQL)          Update existing field values in a table.                                                     
     .Update                Save a recordset.                                                                            
    v                                                                                                                    
      Val                   Extract a numeric value from a string.                                                       
      Var (SQL)             Estimate variance for sample (all records)                                                   
      VarP (SQL)            Estimate variance for population (all records)                                               
      VarType               Return a number indicating the data type of a variable.                                      
    w                                                                                                                    
      Weekday               Return the weekday (1-7) from a date.                                                        
      WeekdayName           Return the day of the week.                                                                  
    y                                                                                                                    
      Year                  Return the year for a given date.                                                            
                                                                                                                         
