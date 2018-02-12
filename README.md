# MoneyPatch
This provides `ofx` and `qif` import patches for Microsoft Money 2005. Indeed, sometimes the files downloaded from banks do not load nicely into MSMoney. The program installs `ofxpatch.bat` and `qifpatch.bat` as filters processing the data and reformatting correctly. 

# Installation
To install, run `ofxpatch.bat` and `qifpatch.bat` in elevated mode. This will register `ofxpatch.bat` and `qifpatch.bat` as the programs to be called when files with `.ofx` and `.qif` extensions are opened.

## Supported banks

Bank | type
---- | ----
Crédit Agricole | ofx
Société Générale | qif


