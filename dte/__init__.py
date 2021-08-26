
from ctypes import CDLL, POINTER, byref, c_bool, c_uint, c_void_p, c_wchar_p

_lib = CDLL(__path__[0] + '/VisualStudio.dll')
_lib.GetBreakPoints.restype = POINTER(c_void_p)
_lib.GetBreakPoint.restype = c_wchar_p
_lib.GetEnabledBreakPoint.restype = c_wchar_p

#--------------------------------------------------------
def open( aVersion ):
  return _lib.Open(aVersion)

#--------------------------------------------------------
def setfile( aPath, aLine ):
  b = _lib.SetFile(aPath, aLine)
  return b

#--------------------------------------------------------
def breakpoints(  ):
  ''' Return tuple of (Breakpoints*, count). '''
  cnt = c_uint()
  bps = _lib.GetBreakPoints(byref(cnt))
  return (bps, cnt.value)

#--------------------------------------------------------
def breakpoint( abps, aIndex ):
  ''' Return tuple of (filename, lineno). '''
  ln = c_uint()
  fl = _lib.GetBreakPoint(abps, aIndex, byref(ln))
  return (fl, ln.value)

#--------------------------------------------------------
def enabledbreakpoint( abps, aIndex, abEnabled ):
  ''' Return tuple of (filename, lineno) of breakpoints that match enabled flag. '''
  ln = c_uint()
  fl = _lib.GetEnabledBreakPoint(abps, aIndex, abEnabled, byref(ln))
  return (fl, ln.value)

#--------------------------------------------------------
def command( aCommand, aParams ):
  return _lib.SendCommand(aCommand, aParams)
