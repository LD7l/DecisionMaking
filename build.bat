@echo off
set CPLEX=C:\Program Files\IBM\ILOG\CPLEX_Studio_Community222
set CPLEXLIB=cplex2220.lib

cl /nologo /EHsc /O2 /MD /DIL_STD /D_CRT_SECURE_NO_WARNINGS main.cpp ^
   /I"%CPLEX%\cplex\include" /I"%CPLEX%\concert\include" ^
   /link /LIBPATH:"%CPLEX%\cplex\lib\x64_windows_msvc14\stat_mda" ^
         /LIBPATH:"%CPLEX%\concert\lib\x64_windows_msvc14\stat_mda" ^
         %CPLEXLIB% ilocplex.lib concert.lib /OUT:roster.exe