lcCurrFolder = addbs(justpath(sys(16)))
set path to fullpath('..\', lcCurrFolder)
_screen.Visible = .F.
do form (lcCurrFolder + 'sample') name loSampleForm
read events
if version(2) = 2
	_screen.Visible = .T.
endif version(2) = 2
