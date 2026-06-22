Filename FILE2025   (  '\\systema44\Migracion\e100\M10125.asu'
                     /*'\\systema44\Migracion\e100\M10225.asu'
                       '\\systema44\Migracion\e100\M10325.asu' 
					   '\\systema44\Migracion\e100\M10425.asu' 
                       '\\systema44\Migracion\e100\M10525.asu' 
					   '\\systema44\Migracion\e100\M10625.asu'
					   '\\systema44\Migracion\e100\M10725.asu'
					   '\\systema44\Migracion\e100\M10825.asu'
					   '\\systema44\Migracion\e100\M10925.asu'
					   '\\systema44\Migracion\e100\M11025.asu'
					   '\\systema44\Migracion\e100\M11125.asu'
					   '\\systema44\Migracion\e100\M11225.asu'*/) LRECL=450;

Filename FILE2026   (  '\\systema44\Migracion\e100\M10126.asu'
                     /*'\\systema44\Migracion\e100\M10226.asu'
                       '\\systema44\Migracion\e100\M10326.asu' 
					   '\\systema44\Migracion\e100\M10426.asu' 
                       '\\systema44\Migracion\e100\M10526.asu' 
					   '\\systema44\Migracion\e100\M10626.asu'
					   '\\systema44\Migracion\e100\M10726.asu'
					   '\\systema44\Migracion\e100\M10826.asu'
					   '\\systema44\Migracion\e100\M10926.asu'
					   '\\systema44\Migracion\e100\M11026.asu'
					   '\\systema44\Migracion\e100\M11126.asu'
					   '\\systema44\Migracion\e100\M11226.asu'*/) LRECL=450;

LIBNAME IMPO '\\DIMPE-D-081\d\COMEX';


%MACRO CIF(VAR);
DATA FILE2025; INFILE FILE2025;LRECL=450;

input 		 @1 FECH $4. @3 MES 2. @5 ADUA 2. @7 PAOR 3. 
             @10 PAPR 3. @13 PACO 3. @16 DEPTO 2. @161 ciudad 29. 
             @18 VIA 2. @20 BAND 3. regimen $ 23-26 
             @27 ACUERDO $3. @30 PBK 13.2 @43 KNETO 13.2 
             @56 CANU 13.2 @69 CODUN $3. @72 POSARA  18. @72 SA 6.
             @90 FOBDO 11.2 @101 FLET 11.2 @112 CIFDO 13.2 
             @125 CIFPE 15. @72 PARTE 4. @72 CAPITULO  2. @305 SEGUROS 13.2 @318 OTROSG 13.2
		     @227 ACTIVID $4. @364 NIT $16. @380 DIGV $1. @381 RAZON $60.;

if posara in (8802400000 8802309000) and regimen in ('C190' 'C196') and nit in ('0000000890912462' '000000890912462' '00000890912462' '0000890912462' '000890912462' '00890912462' '0890912462' '890912462'  
'000000000890100577' '00000000890100577' '0000000890100577' '000000890100577' '00000890100577' '0000890100577' '000890100577' '00890100577' '0890100577' '890100577') then delete;

if regimen='C390' or regimen='C392' or regimen='C393' or regimen='C394' or regimen='C395'
or regimen='C396' or regimen='C397' or regimen='C540' or regimen='C541' or regimen='C545'
or regimen='C546' or regimen='C660' or regimen='C665' or regimen='S100' or regimen='S105'
or regimen='S106' or regimen='S200' or regimen='S140' or regimen='S240' or regimen='C398' 
or regimen='C547'
THEN delete;

DATA FILE2026; INFILE FILE2026;LRECL=450;

input 		 @1 FECH $4. @3 MES 2. @5 ADUA 2. @7 PAOR 3. 
             @10 PAPR 3. @13 PACO 3. @16 DEPTO 2. @161 ciudad 29. 
             @18 VIA 2. @20 BAND 3. regimen $ 23-26 
             @27 ACUERDO $3. @30 PBK 13.2 @43 KNETO 13.2 
             @56 CANU 13.2 @69 CODUN $3. @72 POSARA  18. @72 SA 6.
             @90 FOBDO 11.2 @101 FLET 11.2 @112 CIFDO 13.2 
             @125 CIFPE 15. @72 PARTE 4. @72 CAPITULO  2. @305 SEGUROS 13.2 @318 OTROSG 13.2
		     @227 ACTIVID $4. @364 NIT $16. @380 DIGV $1. @381 RAZON $60.;

if posara in (8802400000 8802309000) and regimen in ('C190' 'C196') and nit in ('0000000890912462' '000000890912462' '00000890912462' '0000890912462' '000890912462' '00890912462' '0890912462' '890912462'  
'000000000890100577' '00000000890100577' '0000000890100577' '000000890100577' '00000890100577' '0000890100577' '000890100577' '00890100577' '0890100577' '890100577') then delete;

if regimen='C390' or regimen='C392' or regimen='C393' or regimen='C394' or regimen='C395'
or regimen='C396' or regimen='C397' or regimen='C540' or regimen='C541' or regimen='C545'
or regimen='C546' or regimen='C660' or regimen='C665' or regimen='S100' or regimen='S105'
or regimen='S106' or regimen='S200' or regimen='S140' or regimen='S240' or regimen='C398' 
or regimen='C547'
THEN delete;

proc sort data=FILE2025;by &var;
proc sort data=FILE2026;by &var;

proc summary data=FILE2025;
var CIFDO CIFPE KNETO PBK CANU;
by &var;
output out=TOTAL2025
sum(CIFDO CIFPE FOBDO KNETO PBK CANU)=CIFDO2025 CIFPE2025 FOBDO2025 KNETO2025 PBK2025 CANU2025;

proc summary data=FILE2026;
var CIFDO CIFPE KNETO PBK CANU;
by &var;
output out=TOTAL2026
sum(CIFDO CIFPE FOBDO KNETO PBK CANU)=CIFDO2026 CIFPE2026 FOBDO2026 KNETO2026 PBK2026 CANU2026;

data IMPO.TOTAL&VAR;
merge TOTAL2026 TOTAL2025;
by &var;
/*if posara = 8805290000;*/
drop _type_;
RUN;
%MEND CIF;

%CIF(posara parte capitulo paor mes);RUN;

PROC SORT DATA=IMPO.totalposara; BY posara;run;

PROC SORT DATA=IMPO.corre62; BY posara; RUN;


DATA impo.EJE; MERGE impo.totalposara (IN=A) IMPO.corre62; BY posara;IF A;
RUN;


PROC EXPORT
 DATA=impo.EJE
 OUTFILE='\\DIMPE-D-081\d\COMEX\IMPO\Resultados\Importaciones 2026.XLSX'
 DBMS=excel
 REPLACE;Run;
