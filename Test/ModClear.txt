Attribute VB_Name = "ModClear"
    Sub Clear_Information()
    
          '------------------------------------------------------------------------
        Rem  Created         : 15/02/11
        Rem  Author          : Leanne Dalgleish
        Rem  Description     : The macro clears the proj-input sheet and re-enters
        'formula referencing D2D sheet
        '
        Rem  Further Info    :
        Rem  Date             Developer     Action/Comments
        '------------------------------------------------------------------------
        '    dd/mm/yyyy       Name          X
        '    dd/mm/yyyy       Name          X
        '------------------------------------------------------------------------
    
        Rem Error handler
    On Error GoTo ErrorHandler
        
    ' Clears Information Page
10  If MsgBox("Do you wish to clear the Proj_input sheet?", vbYesNo) = vbYes Then
         
    Rem Set variables
20  Call modGlobals.Create_Globals
21  strProcName = "Clear_Information"
22  Application.ScreenUpdating = False
23 Application.EnableEvents = False
24 Application.Calculation = xlCalculationManual
25  Sheets("Proj-input").Select

30  With wksProjInput
40     .Range("C3").FormulaR1C1 = "=d2d_policy"
50      .Range("C4").FormulaR1C1 = "=d2d_name"
60      .Range("C5").FormulaR1C1 = "=d2d_dob"
70      .Range("C6").FormulaR1C1 = "=IF(d2d_spouseDOB="""","""",d2d_spouseDOB)"
80      .Range("C7").FormulaR1C1 = "=d2d_nrd"
90      .Range("C8").FormulaR1C1 = "No"
100     .Range("C9").FormulaR1C1 = _
            "=IF(D2D!R[-6]C[59]="""",IF(AND(P2574_Flag=""Yes"",Scheme_Type=""Heritage (Mainframe)"",Protected_NRA=""No"",(INT(((NRD-PRO_DOB)/365))<55)),DATE(YEAR(PRO_DOB)+55,MONTH(PRO_DOB),DAY(PRO_DOB)),NRD),D2D!R[-6]C[59])"
110     .Range("C10").FormulaR1C1 = "=INT(((Revised_NRD-PRO_DOB)/365))"
120     .Range("C11").FormulaR1C1 = _
            "=IF(DATE(YEAR(Calc.date),MONTH(Calc.date),DAY(Calc.date))>=DATE(YEAR(Calc.date),MONTH(PRO_DOB),DAY(PRO_DOB)),SUM((YEAR(Calc.date))-(YEAR(PRO_DOB))),SUM((YEAR(Calc.date))-(YEAR(PRO_DOB))-1))"
130     .Range("C12").FormulaR1C1 = _
            "=IF(DATE(YEAR(Calc.date),MONTH(Calc.date),DAY(Calc.date))<=DATE(YEAR(Calc.date),MONTH(Revised_NRD),DAY(Revised_NRD)),SUM((YEAR(Revised_NRD))-(YEAR(Calc.date))),SUM((YEAR(Revised_NRD))-(YEAR(Calc.date))-1))"
140     .Range("C13").FormulaR1C1 = "=DATEDIF(R[-3]C[4],R[-4]C,""ym"")"
150     .Range("C14").FormulaR1C1 = _
            "=IF(DAY(PRO_DOB+(PRP_Retirement_Age*365.25))<DAY(PRO_DOB),(PRO_DOB+(PRP_Retirement_Age*365.25)+1),IF(DAY(PRO_DOB+(PRP_Retirement_Age*365.25))=DAY(PRO_DOB),(PRO_DOB+(PRP_Retirement_Age*365.25)),(PRO_DOB+(PRP_Retirement_Age*365.25)-1)))"
160     .Range("C15").FormulaR1C1 = _
            "=IF(D2D!R[-12]C[72]<>"""",D2D!R[-12]C[72],DATE(YEAR(Calc.date),MONTH(Calc.date),DAY(Calc.date)))"
170     .Range("C16").FormulaR1C1 = "=d2d_sex"
180     .Range("C17").FormulaR1C1 = "=d2d_maritalSTATUS"
190     .Range("C18").FormulaR1C1 = "N"
200     .Range("C19").FormulaR1C1 = "=d2d_salary"
230     .Range("C21").FormulaR1C1 = "5%"
240     .Range("C22").FormulaR1C1 = "=D2D!R[-19]C[18]"
250     .Range("C23").FormulaR1C1 = "=D2D!R[-20]C[63]"
260     .Range("C24").FormulaR1C1 = "=D2D!R[-21]C[58]"
270     .Range("C25").FormulaR1C1 = "=D2D!R[-22]C[19]"
280     .Range("C26").FormulaR1C1 = "=D2D!R[-23]C[60]"
290     .Range("C27").FormulaR1C1 = "=D2D!R[-24]C[61]"
300     .Range("C28").FormulaR1C1 = "=D2D!R[-25]C[62]"
310     .Range("C29").FormulaR1C1 = "=D2D!R[-26]C[20]"
320     .Range("C30").FormulaR1C1 = "=D2D!R[-27]C[21]"
330     .Range("C31").FormulaR1C1 = "=D2D!R[-28]C[22]"
340     .Range("C32").FormulaR1C1 = "=D2D!R[-29]C[23]"
350     .Range("C33").FormulaR1C1 = "=D2D!R[-30]C[24]"
360     .Range("C34").FormulaR1C1 = "=D2D!R[-31]C[25]"
370     .Range("C35").FormulaR1C1 = "=D2D!R[-32]C[26]"
380     .Range("C36").FormulaR1C1 = "=D2D!R[-33]C[27]"
390     .Range("C37").FormulaR1C1 = "=D2D!R[-34]C[28]"
400     .Range("C38").FormulaR1C1 = "=D2D!R[-35]C[29]"
410     .Range("C39").FormulaR1C1 = "=IF(FVR=0.25%,50000,0)"
    
        ' *** updated by Suman Guria on 15/03/2016 under P3585e ***
'420     .Range("C40").FormulaR1C1 = "=D2D!R[-37]C[30]"
         .Range("D73").FormulaR1C1 = "=D2D!R[-37]C[30]" 'Pradip:P3585e-UAT
'420     .Range("C40").FormulaR1C1 = "=CalcLABPre"      'Pradip:P3585e-UAT
421      .Range("C40").FormulaR1C1 = "=IF(DWPPUP_Flag=""Y"",CalcLABPre,Existing_LABpre)" 'Pradip:P3585e-UAT
        .Range("LABAmount").Value = 0
        ' *** End ***
        
430     .Range("C41").FormulaR1C1 = "=D2D!R[-38]C[57]"
450     .Range("C43").FormulaR1C1 = "=AMC"
460     .Range("C44").FormulaR1C1 = "=Calc.date"
470     .Range("C45").FormulaR1C1 = "0%"
480     .Range("C46").FormulaR1C1 = "=Calc.date"
490     .Range("C47").FormulaR1C1 = "0"
500     .Range("C48").FormulaR1C1 = "=Calc.date-1"
510     .Range("D36").FormulaR1C1 = ""
511     .Range("D29").FormulaR1C1 = ""
520     .Range("D59").FormulaR1C1 = ""

'530     .Range("G3").FormulaR1C1 = "=d2d_nprFV"
'540     .Range("G4").FormulaR1C1 = "=d2d_prpFV"
        'Satya 02.06.2015 - Changed under P3576
530     .Range("G3").FormulaR1C1 = "=d2d_nprFV+NPR_TB-NPR_MVR"
540     .Range("G4").FormulaR1C1 = "=d2d_prpFV+PRP_TB-PRP_MVR"

550     .Range("G5").FormulaR1C1 = "=d2d_contractedOUT"
560     .Range("G6").FormulaR1C1 = "=IF(SUM(D2D!R[-3]C[6]:R[-3]C[9])>0,""Y"",""N"")"
570     .Range("G7").FormulaR1C1 = "Monthly"
580     .Range("G8").FormulaR1C1 = "=d2d_eeFIX+d2d_erFIX"
590     .Range("G9").FormulaR1C1 = "=(d2d_eePERC+d2d_erPERC)/100"
600     .Range("G10").FormulaR1C1 = "=d2d_calcDATE"
610     .Range("G11").FormulaR1C1 = "=d2d_renewal"
620     .Range("G12").FormulaR1C1 = "0%"
630     .Range("G14").FormulaR1C1 = "=IF(((InvGrowth_BasB-0.03)) < 0,InvGrowth_BasB-0.03,(InvGrowth_BasB-0.03))"
631    .Range("R15").FormulaR1C1 = _
        "=IF(((InvGrowth_BasB_Existing-0.03)) < 0,(InvGrowth_BasB_Existing-0.03),(InvGrowth_BasB_Existing-0.03))"
632    .Range("R17").FormulaR1C1 = _
        "=IF(((InvGrowth_BasB_Existing+0.03)) < 0,0,(InvGrowth_BasB_Existing+0.03))"
650     .Range("G16").FormulaR1C1 = "=IF(((InvGrowth_BasB+0.03)) < 0,0,(InvGrowth_BasB+0.03))"
660     .Range("G17").FormulaR1C1 = "2%"
670     .Range("G18").FormulaR1C1 = "4%"
680     .Range("G19").FormulaR1C1 = "6%"
690     .Range("G20").FormulaR1C1 = "=SalGrowth_BasA"
700     .Range("G21").FormulaR1C1 = "=SalGrowth_BasB"
710     .Range("G22").FormulaR1C1 = "=SalGrowth_BasC"
720     .Range("G23").FormulaR1C1 = "1"
730     .Range("G24").FormulaR1C1 = "1"
740     .Range("G25").FormulaR1C1 = "1"
750     .Range("G26").FormulaR1C1 = "=d2d_esc"
760     .Range("G27").FormulaR1C1 = "=d2d_wra"
770     .Range("G28").FormulaR1C1 = "Normal NPR Rate"
780     .Range("G29").FormulaR1C1 = "1"
790     .Range("G30").FormulaR1C1 = "1"
800     .Range("G31").FormulaR1C1 = "1"
810     .Range("G32").FormulaR1C1 = "1"
820     .Range("G33").FormulaR1C1 = "1"
830     .Range("G34").FormulaR1C1 = "Post 1997"
840     .Range("F35:G35").FormulaR1C1 = ""
850     .Range("F36:G38").FormulaR1C1 = ""
860     .Range("F39:G39").FormulaR1C1 = "Please Select the Scheme Format"
870     .Range("F40:G40").FormulaR1C1 = "Please Select a Contract Type"
880     .Range("F41:G41").FormulaR1C1 = "Please Select a Version Number"
890     .Range("F42").FormulaR1C1 = "No"
900     .Range("G42").FormulaR1C1 = "=IF(Waiver_Flag=""Yes"",2.5%,0%)"
910     .Range("A51:A59").FormulaR1C1 = ""
920     .Range("C51:E58").FormulaR1C1 = 0
930     .Range("G51:G58").FormulaR1C1 = 0
940     .Range("C59:C60").FormulaR1C1 = 0
950     .Range("F43").FormulaR1C1 = "No"
960     .Range("G43").FormulaR1C1 = "=NPRAD"
970     .Range("G44").FormulaR1C1 = "NONE"
980     .Range("G45").FormulaR1C1 = "0%"
990     .Range("G46").FormulaR1C1 = "=Calc.date-1"
1000    .Range("G47").FormulaR1C1 = "0"
1010    .Range("B65").FormulaR1C1 = "=IF(LABpre>0,1,"""")"
1020    .Range("B66").FormulaR1C1 = _
            "=IF(R[-1]C="""","""",IF(R[-1]C+1>Term_to_NRD__complete_years+1,"""",R[-1]C+1))"
1030        .Range("C65").FormulaR1C1 = "=LABpre"
1040        .Range("C66").FormulaR1C1 = "=IF(RC[-1]<>"""",R[-1]C,"""")"
1050        .Range("C66").Select
1060        Selection.AutoFill Destination:=.Range("C66:C114")
1070        .Range("C66:C114").Select
1080        .Range("B66").FormulaR1C1 = _
            "=IF(R[-1]C="""","""",IF(R[-1]C+1>Term_to_NRD__complete_years+1,"""",R[-1]C+1))"
1090        .Range("B66").Select
1100        Selection.AutoFill Destination:=.Range("B66:B114")
1110        .Range("B66:B114").Select
1120        .Range("D67").FormulaR1C1 = "65"
1130        .Range("D65").FormulaR1C1 = "No"
1140        .Range("M62").FormulaR1C1 = "No"
1150        .Range("M63").FormulaR1C1 = ""
1160        .Range("M66").FormulaR1C1 = "No"
1170        .Range("M67").FormulaR1C1 = "0%"
1180        .Range("M68").FormulaR1C1 = "0%"
1190        .Range("M69").FormulaR1C1 = "0%"
1200        .Range("M70").FormulaR1C1 = "=LB_AFC_PRP"
1210        .Range("M72").FormulaR1C1 = "No"
1220        .Range("N23").FormulaR1C1 = "Yes"
1230        .Range("N21").FormulaR1C1 = _
            "=IF(Term_to_NRD__complete_years<1,0,IF(Term_to_NRD__complete_years<2,1,IF(Term_to_NRD__complete_years<3,2,3)))"
1240        .Range("C4").Select

1241    .Range("C20").FormulaR1C1 = "=D2D!R[-17]C[80]"
1242    .Range("F42").FormulaR1C1 = "=D2D!R[-39]C[72]"
1243    .Range("G7").FormulaR1C1 = "=D2D!R[-4]C[70]"
1244    .Range("M67").FormulaR1C1 = "=D2D!R[-64]C[66]"
1245    .Range("M68").FormulaR1C1 = "=D2D!R[-65]C[67]"
1246    .Range("M69").FormulaR1C1 = "=D2D!R[-66]C[68]"
1247    .Range("M70").FormulaR1C1 = "=D2D!R[-67]C[69]"
1248    .Range("C51").FormulaR1C1 = "=D2D!R[-48]C[31]"
1249    .Range("C52").FormulaR1C1 = "=D2D!R[-49]C[34]"
1250    .Range("C53").FormulaR1C1 = "=D2D!R[-50]C[37]"
1251    .Range("C54").FormulaR1C1 = "=D2D!R[-51]C[40]"
1252    .Range("C55").FormulaR1C1 = "=D2D!R[-52]C[43]"
1253    .Range("C56").FormulaR1C1 = "=D2D!R[-53]C[46]"
1254    .Range("C57").FormulaR1C1 = "=D2D!R[-54]C[49]"
1255    .Range("C58").FormulaR1C1 = "=D2D!R[-55]C[52]"
1256    .Range("C59").FormulaR1C1 = "=D2D!R[-56]C[55]"
1257    .Range("C60").FormulaR1C1 = "=D2D!R[-57]C[56]"
1258    .Range("D51").FormulaR1C1 = "=D2D!R[-48]C[31]"
1259    .Range("D52").FormulaR1C1 = "=D2D!R[-49]C[34]"
1260    .Range("D53").FormulaR1C1 = "=D2D!R[-50]C[37]"
1261    .Range("D54").FormulaR1C1 = "=D2D!R[-51]C[40]"
1262    .Range("D55").FormulaR1C1 = "=D2D!R[-52]C[43]"
1263    .Range("D56").FormulaR1C1 = "=D2D!R[-53]C[46]"
1264    .Range("D57").FormulaR1C1 = "=D2D!R[-54]C[49]"
1265    .Range("D58").FormulaR1C1 = "=D2D!R[-55]C[52]"
1266    .Range("E51").FormulaR1C1 = "=D2D!R[-48]C[31]"
1267    .Range("E52").FormulaR1C1 = "=D2D!R[-49]C[34]"
1268    .Range("E53").FormulaR1C1 = "=D2D!R[-50]C[37]"
1269    .Range("E54").FormulaR1C1 = "=D2D!R[-51]C[40]"
1270    .Range("E55").FormulaR1C1 = "=D2D!R[-52]C[43]"
1271    .Range("E56").FormulaR1C1 = "=D2D!R[-53]C[46]"
1272    .Range("E57").FormulaR1C1 = "=D2D!R[-54]C[49]"
1273    .Range("E58").FormulaR1C1 = "=D2D!R[-55]C[52]"
1274    .Range("C42").FormulaR1C1 = "=D2D!R[-39]C[73]"
1275    .Range("G51").FormulaR1C1 = "=D2D!R[-48]C[60]"
1276    .Range("G52").FormulaR1C1 = "=D2D!R[-49]C[61]"
1277    .Range("G53").FormulaR1C1 = "=D2D!R[-50]C[62]"
1278    .Range("G54").FormulaR1C1 = "=D2D!R[-51]C[63]"
1279    .Range("G55").FormulaR1C1 = "=D2D!R[-52]C[64]"
1280    .Range("G56").FormulaR1C1 = "=D2D!R[-53]C[65]"
1281    .Range("G57").FormulaR1C1 = "=D2D!R[-54]C[66]"
1282    .Range("G58").FormulaR1C1 = "=D2D!R[-55]C[67]"
1283    .Range("G15").FormulaR1C1 = "=D2D!R[-12]C[77]"
1284    .Range("R16").FormulaR1C1 = "=D2D!R[-13]C[67]"
1285    .Range("R21").FormulaR1C1 = "=InvGrowth_BasB_Existing"
1286    .Range("N12").FormulaR1C1 = "=InvGrowth_BasB"
1287    End With
1310    Application.EnableEvents = True
1320    Application.Calculation = xlCalculationAutomatic
1340    Call modGlobals.Delete_Globals
1350    Application.ScreenUpdating = True
        Rem stops procedure
1355    Sheets("Collect via Macro").Select
1356    Range("C2").ClearContents
1360
        MsgBox ("The Proj_input sheet has been cleared.")

1370 Else: End If
1380 Exit Sub
        Rem Calls the error handler
ErrorHandler:
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
gErr.Number = Err.Number
gErr.Description = Err.Description
gErr.Source = Err.Source
gErr.Erl = Erl
Call modError.Handler
End Sub



