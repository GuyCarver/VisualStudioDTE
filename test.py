
import dte

if not dte.open('VisualStudio.DTE.16.0'):
  print('Error opening DTE')

bps, cnt = dte.breakpoints()
print(bps, cnt)
for i in range(cnt):
  data = dte.breakpoint(bps, i)
  print(data.File, data.Line, data.Enabled)

