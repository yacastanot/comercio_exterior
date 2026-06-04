Filename FILE2025 ( '\\systema44\Migracion\e110\M10125.ava'
                    '\\systema44\Migracion\e110\M10225.ava' 
                    /*'\\systema44\Migracion\e110\M10325.ava'
				    '\\systema44\Migracion\e110\M10425.ava'
				    '\\systema44\Migracion\e110\M10525.ava'
					'\\systema44\Migracion\e110\M10625.ava'
					'\\systema44\Migracion\e110\M10725.ava'
                    '\\systema44\Migracion\e110\M10825.ava'
                    '\\systema44\Migracion\e110\M10925.ava'
					'\\systema44\Migracion\e110\M11025.ava'
					'\\systema44\Migracion\e110\M11125.ava'
					'\\systema44\Migracion\e110\M11225.ava'*/);

Filename FILE2026 ( '\\systema44\Migracion\e110\M10126.ava'
                    '\\systema44\Migracion\e110\M10226.ava' 
                    /*'\\systema44\Migracion\e110\M10326.ava'
				    '\\systema44\Migracion\e110\M10426.ava'
				    '\\systema44\Migracion\e110\M10526.ava'
					'\\systema44\Migracion\e110\M10626.ava'
					'\\systema44\Migracion\e110\M10726.ava'
                    '\\systema44\Migracion\e110\M10826.ava'
                    '\\systema44\Migracion\e110\M10926.ava'
					'\\systema44\Migracion\e110\M11026.ava'
					'\\systema44\Migracion\e110\M11126.ava'
					'\\systema44\Migracion\e110\M11226.ava'*/);

LIBNAME EXPO '\\DIMPE-D-081\d\COMEX';

%macro CALC(VAR);

DATA FILE2025; INFILE FILE2025 LRECL=678;ENCODING=asciiany;

INPUT

@1 FECHPRO 4.	@5 MES 2. @13 adu 2. @42 pais 3. @85 lugsaln 2. @87 lugsal $3. @171 deptor 2.	@139 via 1. @140 BANDERA 3. @143 REGIMEN 1. 
@90 deptpr 2. 	@92 DEX ANTE 14. @144 modalidad 3.	 @147 forpa4 1. @153 posara 10. @153 CAPITULO 2. @48 Ciuexp $20.
@153 PARTE 4. 	@153 SA 6. @173 unidad 2.    @178 cantidad 15.2  @16 Nit $12. @347 Razon $60. @326 DEX $13. 
@208 kneto 15.2  @139 via 1. @238 fobpes 20.2 @223 fobdol 15.2 @37 municipio 5. @273 fletes 15. @320 FECHEMB 6. @339 FECHDEC 8. @288 seguros 15.; 

parte=substr(posara, 1,6);

if POSARA=0901110000 or posara=0901111000 or posara=0901119000 or POSARA=7202600000 then TRA=1;

if 2701<=PARTE<=2704 then TRA=1;
If 2709<=PARTE<=2715 then TRA=1;

IF TRA=. THEN TRA=2;

DATA FILE2026; INFILE FILE2026 LRECL=678; ENCODING=asciiany;

INPUT

@1 FECHPRO 4.	@5 MES 2. @13 adu 2. @42 pais 3. @85 lugsaln 2. @87 lugsal $3. @171 deptor 2.	@139 via 1. @140 BANDERA 3. @143 REGIMEN 1. 
@90 deptpr 2. 	@92 DEX ANTE 14. @144 modalidad 3.	 @147 forpa4 1. @153 posara 10. @153 CAPITULO 2. @48 Ciuexp $20.
@153 PARTE 4. 	@153 SA 6. @173 unidad 2.    @178 cantidad 15.2  @16 Nit $12. @347 Razon $60. @326 DEX $13. 
@208 kneto 15.2  @139 via 1. @238 fobpes 20.2 @223 fobdol 15.2 @37 municipio 5. @273 fletes 15. @320 FECHEMB 6. @339 FECHDEC 8. @288 seguros 15.; 

parte=substr(posara, 1,6);

if POSARA=0901110000 or posara=0901111000 or posara=0901119000 or POSARA=7202600000 then TRA=1;

if 2701<=PARTE<=2704 then TRA=1;
If 2709<=PARTE<=2715 then TRA=1;

IF TRA=. THEN TRA=2;

proc sort data=FILE2025;by &var;
proc sort data=FILE2026;by &var;

proc summary data=FILE2025;
var FOBDOL fobpes;
by &var;
output out=TOTAL2025 sum(FOBDOL KNETO fobpes cantidad)=FOBDOL2025 KNETO2025 fobpes2025 cantidad2025;

proc summary data=FILE2026;
var FOBDOL fobpes;
by &var;
output out=TOTAL2026 sum(FOBDOL KNETO fobpes cantidad)=FOBDOL2026 KNETO2026 fobpes2026 cantidad2026;


data EXPO.TOTAL&VAR;
merge TOTAL2026 TOTAL2025;
by &var;
if posara in (
102299020
201300010
201300090
201300093
201300094
202200000
202300010
202300090
202300093
202300094
202300095/*Carne*/
102299020 /*Ganado vivo*/
)
;
drop _type_;
run;
%mend CALC;


%CALC(posara mes pais modalidad lugsal NIT RAZON); RUN; 


PROC SORT DATA=EXPO.totalposara; BY posara;run;

PROC SORT DATA=EXPO.corre62; BY posara; RUN;


DATA expo.EJE; MERGE expo.totalposara (IN=A) EXPO.corre62; BY posara;IF A;
RUN;


PROC EXPORT
 DATA=Expo.Eje
 OUTFILE='\\DIMPE-D-081\d\COMEX\EXPO\Resultados\Exportaciones GanadovsCarne2025-2026_NIT.XLSX'
 DBMS=excel
 REPLACE;Run;