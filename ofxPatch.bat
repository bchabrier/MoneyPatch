@ECHO OFF
SETLOCAL ENABLEDELAYEDEXPANSION

: define mapping of tiers, separated by :
: * is supported at the beginning and end of a tier
set tiers=toto=titi:

: linefeed - keep the 2 empty lines!
set LF=^


: leave 2 previous lines blank

rem Find MSMoney path
rem
rem trouve la parenthese dans: (par défaut)   REG_SZ    D:\Program Files\Microsoft Money 2005\MNYCoreFiles\MSMoney.exe -url:%1
rem                                       |<----- %%a ----
for /f "tokens=1,* delims=)" %%a in ('reg query HKEY_CLASSES_ROOT\money\Shell\Open\Command /ve') do (
    set l=%%b
    rem trouve la fin dans:   REG_SZ    D:\Program Files\Microsoft Money 2005\MNYCoreFiles\MSMoney.exe -url:%1
    rem                                |<----- %%b ----
    for /f "tokens=1,* delims= " %%a in ('echo.!l!') do (
        set moneyPath=%%~dpb
    )
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
            if "%1"=="" pause
        )
    )
)

set file=%1
if "%file%"=="" exit /B

set ofile=%TEMP%\%~nx1

echo Creating %ofile%...
(
    for /f "tokens=1 delims==" %%a in ('type "%file%"') do @(
        set line=%%a
        set str=@@@!line!

        set key=^<NAME^>
        call set "val=%%str:@@@!key!=%%"
        if not "!str!"=="!val!" (
            rem <NAME> line
            set nameval=!val!

            : parse the tiers
            for %%A in ("!LF!") do (
                set newname=
                for /f "eol== tokens=1,2 delims==" %%B in ("!tiers::=%%~A!") do (
                    if "!newname!"=="" (
                        : echo %%B -^> %%C >&2
                        set re=%%B
                        if "!re:~-1!"=="*" (
                            : ending with star, we will compare with the beginning only
                            set re=!re:~0,-1!
                        ) else (
                            set re=!re!@@@@
                        )
                        if "!re:~0,1!"=="*" (
                            : starting  with star, we will compare with the end only
                            set re=!re:~1!
                        ) else (
                            set re=@@@@!re!
                        )
                        : echo re=!re! >&2
                        set search=@@@@!nameval!@@@@
                        call set "r=%%search:!re!=-%%"
                        if not "!search!"=="!r!" (
                            : string was found in tier name
                            : echo !nameval!, !r! >&2
                            : so if the replacement is exactly "-"" then we substitute
                            set newname=%%C
                        )
                    )
                )
                if not "!newname!"=="" (
                    echo Tier !nameval! -^> !newname! >&2
                    rem echo ^<NAME^>!newname!
                    set nameval=!newname!
                ) else (
                    rem echo Tier !nameval! ignored >&2                          
                )
            )
        ) else (
            set key=^<MEMO^>
            call set "val=%%str:@@@!key!=%%"
            if not "!str!"=="!val!" (
                rem <MEMO> line
                set memo=!line! !nameval!
                if "!val!"=="CHEQUE EMIS" (
                    rem take first number before / (sometimes got 12345/00000/0000)
                    for /f "tokens=1 delims=/" %%A in ("!nameval!") do echo ^<CHECKNUM^>%%A
                ) else (
                        if not "!nameval!"=="adsfadsf" echo ^<NAME^>!nameval!
                )
                echo !memo!
            ) else (
                set nameval=
                echo.!line!
            )
        )
    ) 
) > %ofile%
echo.

more %ofile%

<nul set /p =Importing file into Money...
"%moneyPath%mnyimprt.exe" %ofile%
echo.
rem pause
