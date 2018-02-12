@ECHO OFF
SETLOCAL ENABLEDELAYEDEXPANSION

rem list of card patterns to be found in the memo field (starting with N), separated by :
set cards=CARTE X1916:

: linefeed - keep the 2 empty lines!
set LF=^


: leave 2 previous lines blank

set file=%1
if not "%2"=="" (
    set file="!file! %2"
)
: remove surrounding " if any
for /f "useback tokens=*" %%a in ('%file%') do set file=%%~a

rem Find MSMoney path
rem
rem trouve la parenthese dans: (par d�faut)   REG_SZ    D:\Program Files\Microsoft Money 2005\MNYCoreFiles\MSMoney.exe -url:%1
rem                                       |<----- %%a ----
for /f "tokens=1,* delims=)" %%a in ('reg query HKEY_CLASSES_ROOT\money\Shell\Open\Command /ve') do (
    set l=%%b
    rem trouve la fin dans:   REG_SZ    D:\Program Files\Microsoft Money 2005\MNYCoreFiles\MSMoney.exe -url:%1
    rem                                |<----- %%b ----
    for /f "tokens=1,* delims= " %%a in ('echo.!l!') do (
        set moneyPath=%%~dpb
    )
)

rem Find moneyFile path
rem
for /f "tokens=2,*" %%a in ('reg query HKEY_CURRENT_USER\Software\Microsoft\Money\14.0 /v CurrentFile') do (
    set iniFile=MoneyPatch.ini
)

set openCmd=\"%moneyPath%%~nx0\" %%1

rem b is the RHS member compared
set b=%openCmd:\"=%
set openCmdEsc=%openCmd:)=^)%
set thisProgEsc=%0
set thisProgEsc=%thisProgEsc:)=^)%
set ext=%~n0
set ext=%ext:~0,3%

rem check if this program is already registered
set errf="%TEMP%\err"
for /f "tokens=1,* delims=)" %%a in ('reg query HKEY_CLASSES_ROOT\%ext%.Document\Shell\Open\Command /ve') do (
    set l=%%b
    for /f "tokens=1,* delims= " %%a in ('echo.!l!') do (
        set l=%%b
        set a=!l:"=!
        rem echo "a=!a!" & echo "b=!b!"
        fc "%~0" "%moneyPath%%~nx0" >NUL
        if not "!a!!errorlevel!"=="!b!0" (
            echo The program needs to be installed, trying to install...
            rem register this program
            rem
            reg add HKEY_CLASSES_ROOT\%ext%.Document\Shell\Open\Command /ve /d "%openCmdEsc%" /f>NUL 2>%errf%  
            copy "%thisProgEsc%" "%moneyPath%" >NUL 2>>%errf%
            for /f "tokens=*" %%a in ('type %errf%') do set err=%%a
            if not "!err!"=="" (
                type %errf%
                echo Please run as administrator to install the program
                pause
                del %errf%
                exit /B
            )
            del %errf%
            echo Program successfully installed.
            if "%file%"=="" pause
        )
    )
)


if "%file%"=="" exit /B
set accounts=MAIN
for %%A in ("!LF!") do (
    for /f "tokens=1 delims=:" %%a in ("!cards::=%%~A!") do (
        set regexp=M.*%%a
        rem blanks in strings mean "or", so replace them with .
        set regexp=!regexp: =.!
        findstr /B "!regexp!" "%file%" >nul
        if !errorlevel!==0 (
            set accounts=!accounts!:%%a
        )
    )
)

for %%A in ("!LF!") do (
    for /f "tokens=* delims=:" %%a in ("!accounts::=%%~A!") do (
        set account=%%a
        set memo=M.*!account!

        : remove () part in case the file is a copy
        set ofile=%~n1
        for /f "tokens=1 delims=(" %%a in ("!ofile!") do set n1=%%~a
        set ofile=%TEMP%\!n1!_!account: =_!.ofx

        <nul set /p =Creating !ofile!...
        (
            echo OFXHEADER:100
            echo DATA:OFXSGML
            echo VERSION:102
            echo SECURITY:NONE
            echo ENCODING:USASCII
            echo CHARSET:1252
            echo COMPRESSION:NONE
            echo OLDFILEUID:NONE
            echo NEWFILEUID:NONE
            echo ^<OFX^>
            echo ^<SIGNONMSGSRSV1^>
            echo ^<SONRS^>
            echo ^<STATUS^>
            echo ^<CODE^>0
            echo ^<SEVERITY^>INFO
            echo ^</STATUS^>
            echo ^<DTSERVER^>20180127204948
            echo ^<LANGUAGE^>FRA
            echo ^</SONRS^>
            echo ^</SIGNONMSGSRSV1^>
            echo ^<BANKMSGSRSV1^>
            echo ^<STMTTRNRS^>
            echo ^<TRNUID^>!n1!
            echo ^<STATUS^>
            echo ^<CODE^>0
            echo ^<SEVERITY^>INFO
            echo ^</STATUS^>

            echo ^<STMTRS^>
            echo ^<CURDEF^>EUR
            echo ^<BANKACCTFROM^>
        rem    echo ^<BANKID^>19106
        rem    echo ^<BRANCHID^>00645
            echo ^<BANKID^>00000
            echo ^<BRANCHID^>00000
            if "!account!"=="MAIN" (
                echo ^<ACCTID^>!n1!
            ) else (
                echo ^<ACCTID^>!account!               
            )
            echo ^<ACCTTYPE^>CHECKING
            echo ^</BANKACCTFROM^>

            echo ^<BANKTRANLIST^>
            echo ^<DTSTART^>20180101000000
            echo ^<DTEND^>20180126235959

            set /a count=0
            set /a trcount=0
            for /f "tokens=1 delims==" %%a in ('type "%file%"') do @(
                set line=%%a
                set firstchar=!line:~0,1!

                if "!firstchar!"=="D" (
                    set FD=!line:~1!
                ) else (
                    if "!firstchar!"=="T" (
                        set FT=!line:~1!
                    ) else (
                        if "!firstchar!"=="N" (
                            set FN=!line:~1!
                        ) else (
                            if "!firstchar!"=="P" (
                                set FP=!line:~1!
                            ) else (
                                if "!firstchar!"=="M" (
                                    set FM=!line:~1!
                                ) else (
                                    if "!firstchar!"=="^" (
                                        rem log transaction

                                        rem D17/01/2018
                                        rem T-16.10
                                        rem NCarte
                                        rem PCOTISATION JAZZ
                                        rem MCOTISATION JAZZ  

                                        if not "!count!"=="-10" ( rem used to limit the number of transactions for debug

                                            set /a count="!count! + 1"

                                            rem is it a credit card ?
                                            set found=0
                                            for /f "tokens=* delims=:" %%a in ("!cards::=%%~A!") do (
                                                set card=%%a
                                                if not !found!==1 (
                                                    call set repl=%%FM:!card!=%%
                                                    if not "!repl!"=="!FM!" (
                                                        rem found the credit card in the Memo
                                                        set found=1
                                                        set CC=!card!
                                                        rem echo found CC:!CC!>&2
                                                    )
                                                )
                                            )

                                            rem DAB are stored in main account
                                            if not "!FM:RETRAIT DAB=!"=="!FM!" (
                                                set found=0
                                            )

                                            set do=0
                                            if !found!==1 (
                                                rem is a credit card
                                                if "!account!"=="!CC!" (
                                                    rem the right CC, so let's log the transaction, with no name
                                                    set do=1
                                                    set FN=
                                                )
                                            ) else (
                                                rem no CC, log if main
                                                if "!account!"=="MAIN" set do=1
                                            )

                                            if "!FN!"=="Carte" set FN=
                                            if "!FN!"=="Pr�lvmt" set FN=
                                            if "!FN!"=="Pr�lvmt" set FN=

                                            if !found!_!do!==1_1 (
                                                rem special treatment for VISA CARD
                                                set beg=@@@!FM!
                                                set datepos=0
                                                rem CARTE X1916 REMBT 28/11 LEROY MERLIN
                                                set beg=!beg:@@@CARTE X1916 REMBT=!
                                                rem CARTE X8969 RETRAIT DAB SG 29/12 09H27 LE CANNET ROCHEVILLE 00910790
                                                set beg=!beg:@@@CARTE X1916 RETRAIT DAB SG=!
                                                rem CARTE X1916 03/11 AVENANCE ENTREPR COMMERCE ELECTRONIQUE
                                                set beg=!beg:@@@CARTE X1916=!

                                                if "!beg:@@@=!"=="!beg!" (
                                                    rem replacement was done, so we should have the date in pos 1
                                                    for /F "tokens=1,*" %%a in ("!beg!") do (
                                                        set FN=%%b
                                                        set date=%%a
                                                        set year=!FD:~6,4!
                                                        set /A diff=1!FD:~3,2!!FD:~0,2!-1!date:~3,2!!date:~0,2!
                                                        if "!diff:~0,1!"=="-" set /A year=!year!-1
                                                        set FD=!date!/!year!
                                                        rem echo FDapres=!FD! >&2
                                                        rem echo FN=!FN! >&2
                                                    )
                                                )
                                            )

                                            rem case of goldcard
                                            rem 2201/FORVILLE CANNES 06 CANNES
                                            if "!FM:~4,1!"=="/" (
                                                set numbers="str01020304050607080910111213141516171819202122232425262728293031"
                                                call set numb=%%numbers:!FM:~0,2!=%%
                                                if not "!numb!"=="!numbers!" (
                                                    : 2 first chars are a number
                                                    call set numb=%%numbers:!FM:~2,2!=%%
                                                    if not "!numb!"=="!numbers!" (
                                                        : 2 next chars are a number
                                                        set FN=!FM:~5!
                                                        set date=!FM:~0,2!/!FM:~2,2!
                                                        set year=!FD:~6,4!
                                                        set /A diff=1!FD:~3,2!!FD:~0,2!-1!date:~3,2!!date:~0,2!
                                                        if "!diff:~0,1!"=="-" set /A year=!year!-1
                                                        set FD=!date!/!year!
                                                    )
                                                )
                                            )

                                            rem ignore transactions before 2018
                                            set /a diffy=!FD:~6,4!-2018
                                            if "!diffy:~0,1!"=="-" set do=0

                                            if "!do!"=="1" (
                                                set /a trcount=!trcount!+1
                                                echo ^<STMTTRN^>
                                                echo ^<TRNTYPE^>OTHER
                                                echo ^<DTPOSTED^>!FD:~6,4!!FD:~3,2!!FD:~0,2!
                                                echo ^<TRNAMT^>!FT!
                                                echo ^<FITID^>!FD:~6,4!!FD:~3,2!!FD:~0,2!!COUNT!
                                                if not "!FM:CHEQUE=!"=="!FM!" (
                                                    rem check, take the first element after CHEQUE
                                                    for /F "tokens=2" %%a in ("!FM!") do echo ^<CHECKNUM^>%%a
                                                ) else (
                                                    rem Not a check
                                                    if "!FN!"=="" (
                                                        echo ^<NAME^>^</NAME^>
                                                    ) else (
                                                        rem careful, only 64 chars are supported in NAME!!!
                                                        echo ^<NAME^>!FN:~0,64!
                                                    )
                                                )
                                                echo ^<MEMO^>!FM!
                                                echo ^</STMTTRN^>
                                            )

                                            rem reset fields
                                            set FD=
                                            set FT=
                                            set FN=
                                            set FP=
                                            set FM=
                                        )
                                    ) else (
                                        set unknown=1
                                        if "!line!"=="Bank" (
                                            rem do nothing
                                            set unknown=0
                                        ) 
                                        if "!line!"=="CCard" (
                                            rem do nothing
                                            set unknown=0
                                        ) 
                                        if !unknown!==1 echo Unknown statement: !line! >&2
                                    )
                                )
                            )
                        )
                    )
                )
            ) 

            echo ^</BANKTRANLIST^>
            echo ^<LEDGERBAL^>
            echo ^<BALAMT^>0
            echo ^<DTASOF^>20000101
            echo ^</LEDGERBAL^>
            echo ^<AVAILBAL^>
            echo ^<BALAMT^>0
            echo ^<DTASOF^>20000101
            echo ^</AVAILBAL^>
            echo ^</STMTRS^>
            echo ^</STMTTRNRS^>
            echo ^</BANKMSGSRSV1^>
            echo ^</OFX^>

        ) > !ofile!
        echo. !trcount! transactions

        rem more "!ofile!"

        <nul set /p =Importing file into Money...
        "%moneyPath%mnyimprt.exe" !ofile!
        echo.
    )
)
pause



