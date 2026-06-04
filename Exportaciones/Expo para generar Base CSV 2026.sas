FILENAME  DATOS    ( '\\systema44\Migracion\E110\M10126.AVA '
                     '\\systema44\Migracion\E110\M10226.AVA '
                     /*'\\systema44\Migracion\E110\M10326.AVA '
				     '\\systema44\Migracion\E110\M10426.AVA '
                     '\\systema44\Migracion\E110\M10526.AVA '
					 '\\systema44\Migracion\E110\M10626.AVA '
                     '\\systema44\Migracion\E110\M10726.AVA '
                     '\\systema44\Migracion\E110\M10826.AVA '
          			 '\\systema44\Migracion\E110\M10926.AVA '
					 '\\systema44\Migracion\E110\M11026.AVA '
                     '\\systema44\Migracion\E110\M11126.AVA '
					 '\\systema44\Migracion\E110\M11226.AVA '*/)LRECL=678;

DATA COMER;
        INFILE DATOS ;
         INPUT
@1  ANO_PROC      4.     @5  MES_PROC      2. 
@7  NRO_ART       4.     @11 COD_INCO      2.
@13 COD_ADUA      2.     @15 TIP_IDENT     1.   
@16 NIT         $ 12.    @28 TIP_USUA     2.    
@30 COD_USUA      5.     @35 TIP_EXPO      2.
@37 DPTOMPIO      5.     @42 COD_PAIS     $ 3.   
@45 COD_PAI1     $ 3.    @48 CIU_PAIS    $ 20.  
@68 AUT_EMB     $ 16.    @84 TIP_DECL     1.  
@85 COD_SAL1      2.     @87 COD_SALI     $ 3.    
@90 DPTO_PRO      2.     @92 DEC_EXP     $ 14.  
@106 ANO_EXP      4.     @110 MES_EXP      2.    
@112 DEC_IMPO   $ 14.    @126 ANO_IMPO    4.  
@130 MES_IMPO     2.     @132 COD_MIMP    $ 4.
@136 COD_MONE    $ 3.    @139 COD_VIA     1.  
@140 BANDERA      3.     @143 COD_REGI     1.   
@144 COD_MODA    $ 3.    @147 FOR_PAGO    1.
@148 COD_EMB     $ 1.    @149 COD_DAT    $ 1. 
@150 CER_ORIG     1.     @151 SIS_ESPE     1.    
@152 EXP_TRAN    $ 1.    @153 POS_ARAN    10.
@163 N_SUPLEM   $ 4.     @167 N_COMPLE   $ 4. 
@171 DPTO_ORI     2.     @173 COD_MEDI     2.
@175 CODUNI2     $ 3.    @178 CANT_UNI    15.2  
@193 KILO_BR3    15.2    @208 KILO_NET    15.2
@223 VAL_FOB     15.2    @238 FOB_PES3    20.2 
@258 VAL_ANAL    15.2    @273 FLETES3     15.2  
@288 SEGURO3     15.2    @303 OTROSG3     15.2
@318 COD_AEMB     2.     @320 ANO_EMB      4.   
@324 MES_EMB      2.     @326 NUM_DECL   $ 13. 
@339 ANO_DEC      4.     @343 MES_DEC      2.
@345 DIA_DEC      2.     @347 RACIAL     $ 60.
@407 DIR_EXPO   $ 60.    @467 IDENTIFT   $ 12.
@479 NOM_APEL   $ 60.    @539 RACIALIM   $ 60.
@599 DIR_PDES   $ 80.  	 @153 parte    4.;


IF COD_MODA IN('001','003','105','301','302','303','304','305','306','307','399','308','309','701','500','202')THEN DELETE;
IF COD_MODA IN( '401','402','403' )AND FOR_PAGO = 2 THEN DELETE; 
IF POS_ARAN IN(7108200000,4907002000,4907003000) THEN DELETE;
IF  NIT IN('890912462','0890912462','00890912462','000890912462','8909124622','08909124622','008909124622') 
AND COD_MODA IN('401','402') AND  POS_ARAN IN (8802400000,8802309000) THEN DELETE;
IF  NIT IN('890100577','0890100577','00890100577','000890100577','8901005776','08901005776','008901005776')
AND COD_MODA IN('401','402') AND  POS_ARAN IN (8802400000,8802309000) THEN DELETE;

 
RUN;
