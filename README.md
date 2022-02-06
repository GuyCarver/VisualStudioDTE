# VisualStudioDTE
Visual Studio 2019 DTE dll for use with sublime text 4 plugin.

The dll may also be used by other c++ or python apps.

dte/__init__.py is the python interface to the .dll

## Functions are:
  ### bool Open( const wchar_t *apVersion )
    Opens a connection to an active instance of visual studio.
    apVersion format is "VisualStudio.DTE.??.0" where ?? = 15 for VS2017, 16 for VS2019 and 17 for VS2022
    Returns true if opening the VisualStudio connection worked.
  
  ### void Close()
    Closes the connection
    
  ### bool SetFile( const wchar_t *apFileName, uint32_t aiLineNumber )
    Open the given file and set the cursor at the given line.  The file name must be a full path.
    
  ### void *GetBreakPoints( uint32_t &aCount )
    Get an array of breakpoints. Void* points to the data and aCount will hold the number of entries.
    Access to a specific breakpoint is handled with GetBreakPoint()
    
  ### const BreakPointData GetBreakPoint( void *apBreaks, uint32_t aIndex )
    Get the BreakPointData (see VisualStudio.h) from the apBreaks data from the given index.
    The data is return by GetBreakPoints() and the index must be < the aCount value returned by GetBreakPoints().
    
  ### bool SendCommand( const wchar_t *apCommand, const wchar_t *apArgs )
    Send a command to Visual Studio.
    Parameter examples are:
      ("File.OpenFile", filename)
      ("Build.Compile", "")
      ("Debug.ToggleBreakPoint", "")
      
  ### int32_t AddOutputWindow( const wchar_t *apWindowName )
    Adds a pane to the OutputWindow tab with the given name. Additional calls with the same name will return the same index.
    The returned 1 based index (0 = no pane) is used to send command to the pane.
    
  ### void OutputToPane( int32_t aIndex, const wchar_t *apString )
    Output given string to the OutputWindow pane at the given index (returned by AddOutputWindow()).
  ### void ClearPane( int32_t aIndex )
    Clear output of the OutputWindow pane at the given index (returned by AddOutputWindow()).

  ### void ActivatePane( int32_t aIndex )
    Activer the OutputWindow pane at the given index (returned by AddOutputWindow()).




  

