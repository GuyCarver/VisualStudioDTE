
from ctypes import CDLL, POINTER, Structure, byref, c_bool, c_uint, c_void_p, c_wchar_p

#Structure containing breakpoint data returned by breakpoint()
class BreakPointData(Structure):
  _fields_ = [('File', c_wchar_p), ('Line', c_uint), ('Enabled', c_bool)]

_lib = CDLL(__path__[0] + '/VisualStudio.dll')

_lib.GetBreakPoints.restype = POINTER(c_void_p)
_lib.GetBreakPoint.restype = BreakPointData

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
def breakpoint( apbs, aIndex ):
  ''' Return BreakPointData structure with data for breakpoint at given index. '''
  return _lib.GetBreakPoint(apbs, aIndex)

#--------------------------------------------------------
def command( aCommand, aParams ):
  return _lib.SendCommand(aCommand, aParams)
