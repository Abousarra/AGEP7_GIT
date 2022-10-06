Attribute VB_Name = "Module1"
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public mCapHwnd As Long
Public Const CONNECT As Long = 1034
Public Const DISCONNECT As Long = 1035
Public Const GetObject As Long = 1036
Public Const COPY As Long = 1054
Public PicFile As String
Public xe As Double
Public xc As String
Public xs As String
Public vtx1 As String
Public vtx2 As String
Public base As String
Public Adat As String
Public Aheu As String
Public Atyp As String
Public Amon As String
Public Adet As String
Public Acom As String
Public Auti As String
Public co As ADODB.Connection
Public co2 As ADODB.Connection
Public et2 As ADODB.Recordset
Public pp2 As ADODB.Recordset
Public fc5 As ADODB.Recordset
Public co3 As ADODB.Connection
Public et3 As ADODB.Recordset
Public ce3 As ADODB.Recordset
Public pr3 As ADODB.Recordset
Public pp3 As ADODB.Recordset
Public fc3 As ADODB.Recordset
Public pf3 As ADODB.Recordset
Public ut As ADODB.Recordset
Public ou As ADODB.Recordset
Public pr As ADODB.Recordset
Public eb As ADODB.Recordset
Public sr As ADODB.Recordset
Public an As ADODB.Recordset
Public fc As ADODB.Recordset
Public pf As ADODB.Recordset
Public cl As ADODB.Recordset
Public cr As ADODB.Recordset
Public et As ADODB.Recordset
Public ie As ADODB.Recordset
Public cf2 As ADODB.Recordset
Public cf1 As ADODB.Recordset
Public mt As ADODB.Recordset
Public nt As ADODB.Recordset
Public pe As ADODB.Recordset
Public pp As ADODB.Recordset
Public pc As ADODB.Recordset
Public em As ADODB.Recordset
Public pl As ADODB.Recordset
Public cp As ADODB.Recordset
Public cf As ADODB.Recordset
Public cs As ADODB.Recordset
Public ct As ADODB.Recordset
Public dp As ADODB.Recordset
Public bn As ADODB.Recordset
Public ca As ADODB.Recordset
Public rg As ADODB.Recordset
Public pd As ADODB.Recordset
Public Enum Ahmede
    arabic = vbMsgBoxRight + vbMsgBoxRtlReading
End Enum

Function cont()
Set co = New ADODB.Connection
Set ut = New ADODB.Recordset
Set ou = New ADODB.Recordset
Set pr = New ADODB.Recordset
Set eb = New ADODB.Recordset
Set sr = New ADODB.Recordset
Set an = New ADODB.Recordset
Set fc = New ADODB.Recordset
Set pf = New ADODB.Recordset
Set cl = New ADODB.Recordset
Set cr = New ADODB.Recordset
Set et = New ADODB.Recordset
Set ie = New ADODB.Recordset
Set cf2 = New ADODB.Recordset
Set cf1 = New ADODB.Recordset
Set mt = New ADODB.Recordset
Set nt = New ADODB.Recordset
Set pe = New ADODB.Recordset
Set pp = New ADODB.Recordset
Set pc = New ADODB.Recordset
Set em = New ADODB.Recordset
Set pl = New ADODB.Recordset
Set cp = New ADODB.Recordset
Set cf = New ADODB.Recordset
Set cs = New ADODB.Recordset
Set ct = New ADODB.Recordset
Set dp = New ADODB.Recordset
Set bn = New ADODB.Recordset
Set ca = New ADODB.Recordset
Set rg = New ADODB.Recordset
Set pd = New ADODB.Recordset
If Start_UP.Label1.Caption <> "" Then
base = Start_UP.Label1.Caption
'ane = "2012-2013"
Else
base = Interface.SBB1.Panels(1).Text
End If
'base = Interface.SBB1.Panels(1).Text
co.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co.ConnectionString = App.Path & "\" & base & ".mdb"
co.Open
ut.Open "select*from Utilisateurs", co, adOpenKeyset, adLockOptimistic
ou.Open "select*from Options_uti", co, adOpenKeyset, adLockOptimistic
pr.Open "select*from Partennaires", co, adOpenKeyset, adLockOptimistic
eb.Open "select*from Etablissement", co, adOpenKeyset, adLockOptimistic
sr.Open "select*from Series", co, adOpenKeyset, adLockOptimistic
an.Open "select*from annees", co, adOpenKeyset, adLockOptimistic
fc.Open "select*from Fonctionnaires", co, adOpenKeyset, adLockOptimistic
pf.Open "select*from Professeurs", co, adOpenKeyset, adLockOptimistic
cl.Open "select*from classes", co, adOpenKeyset, adLockOptimistic
cr.Open "select*from Correspondants", co, adOpenKeyset, adLockOptimistic
et.Open "select*from Etudiants order by cla,num ASC", co, adOpenKeyset, adLockOptimistic
ie.Open "select*from IM_Noms_Etudiants", co, adOpenKeyset, adLockOptimistic
cf2.Open "select*from Coffdevoirs", co, adOpenKeyset, adLockOptimistic
cf1.Open "select*from Coffdevoirs1", co, adOpenKeyset, adLockOptimistic
mt.Open "select*from Matieres order by lng,mat ASC", co, adOpenKeyset, adLockOptimistic
nt.Open "select*from Notes", co, adOpenKeyset, adLockOptimistic
pe.Open "select*from Pointage_E order by dat DESC", co, adOpenKeyset, adLockOptimistic
pp.Open "select*from Pointage_P", co, adOpenKeyset, adLockOptimistic
pc.Open "select*from Pourcentage order by moi,cla ASC", co, adOpenKeyset, adLockOptimistic
em.Open "select*from Emplois", co, adOpenKeyset, adLockOptimistic
pl.Open "select*from Pointage_C", co, adOpenKeyset, adLockOptimistic
cp.Open "select*from Caisse_PRT", co, adOpenKeyset, adLockOptimistic
cf.Open "select*from Caisse_FNC", co, adOpenKeyset, adLockOptimistic
cs.Open "select*from Caisse_PRF", co, adOpenKeyset, adLockOptimistic
ct.Open "select*from Caisse_ETU order by sri ASC", co, adOpenKeyset, adLockOptimistic
dp.Open "select*from Caisse_DPS", co, adOpenKeyset, adLockOptimistic
bn.Open "select*from Caisse_BNK", co, adOpenKeyset, adLockOptimistic
ca.Open "select*from Caisse_ARC", co, adOpenKeyset, adLockOptimistic
rg.Open "select*from Notes", co, adOpenKeyset, adLockOptimistic
pd.Open "select*from pointage_date", co, adOpenKeyset, adLockOptimistic
End Function
Public Sub unloadforms()
Unload Utilisateurs
Unload Partenaires
Unload Etablissement
Unload Fonctionnaires
Unload Professeurs
Unload Classes
Unload Correspondants
Unload Etudiants
Unload Cartes
Unload Matieres
Unload Notes
Unload Notes_E
Unload Notes_C
Unload Notes_B
Unload Pointage_E
Unload Pointage_P
Unload Emplois
Unload Caisse_FNC
Unload Caisse_PRT
Unload Caisse_PRF
Unload Caisse_ETU
Unload Caisse_DPS
Unload Caisse_SLD
Unload Caisse_BNK
Unload Compte_TRS
Unload Compte_ARC
Unload Compte_PRT
Unload Compte_FNC
Unload Compte_PRF
Unload Compte_ETU
Unload Compte_DPS
Unload Compte_BNK
Unload Compte_CLS
Unload Archives_AS
Unload Coin_CRS
Unload Recherches
End Sub
Public Sub Series()
xc = xe
If xe < 10 Then
xs = "000000" + xc
ElseIf xe < 100 And xe >= 10 Then
xs = "00000" + xc
ElseIf xe < 1000 And xe >= 100 Then
xs = "0000" + xc
ElseIf xe < 10000 And xe >= 1000 Then
xs = "000" + xc
ElseIf xe < 100000 And xe >= 10000 Then
xs = "00" + xc
ElseIf xe < 1000000 And xe >= 100000 Then
xs = "0" + xc
Else
xs = xc
End If
End Sub
Function cont2()
Set co2 = New ADODB.Connection
Set et2 = New ADODB.Recordset
Set pp2 = New ADODB.Recordset
Set fc5 = New ADODB.Recordset
base = Interface.SBB1.Panels(1).Text
co2.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co2.ConnectionString = App.Path & "\" & base & ".mdb"
co2.Open
et2.Open "select*from Etudiants order by cla,num ASC", co2, adOpenKeyset, adLockOptimistic
pp2.Open "select*from Pointage_P", co2, adOpenKeyset, adLockOptimistic
fc5.Open "select*from Fonctionnaires", co2, adOpenKeyset, adLockOptimistic
End Function
Function cont3()
Set co3 = New ADODB.Connection
Set et3 = New ADODB.Recordset
Set ce3 = New ADODB.Recordset
Set pr3 = New ADODB.Recordset
Set pp3 = New ADODB.Recordset
Set fc3 = New ADODB.Recordset
Set pf3 = New ADODB.Recordset
base = Interface.SBB1.Panels(1).Text
co3.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
co3.ConnectionString = App.Path & "\R2016-2017.mdb"
co3.Open
et3.Open "select*from Tetudiants", co3, adOpenKeyset, adLockOptimistic
ce3.Open "select*from Tcompteetudiants order by ser ASC", co3, adOpenKeyset, adLockOptimistic
pr3.Open "select*from Tprofesseurs", co3, adOpenKeyset, adLockOptimistic
pp3.Open "select*from Tpresences", co3, adOpenKeyset, adLockOptimistic
fc3.Open "select*from Tfonctionnaires", co3, adOpenKeyset, adLockOptimistic
pf3.Open "select*from Tpayfonctionnaires", co3, adOpenKeyset, adLockOptimistic
End Function

Public Sub archive_caisse()
ca.AddNew
ca!dat = Adat
ca!heu = Aheu
ca!typ = Atyp
ca!mon = Amon
ca!det = Adet
ca!com = Acom
ca!uti = Auti
ca.Update

End Sub

Public Sub verif_n_serie()
If Len(vtx1) = 1 Then
vtx2 = "000000" + vtx1
ElseIf Len(vtx1) = 2 Then
vtx2 = "00000" + vtx1
ElseIf Len(vtx1) = 3 Then
vtx2 = "0000" + vtx1
ElseIf Len(vtx1) = 4 Then
vtx2 = "000" + vtx1
ElseIf Len(vtx1) = 5 Then
vtx2 = "00" + vtx1
ElseIf Len(vtx1) = 6 Then
vtx2 = "0" + vtx1
ElseIf Len(vtx1) = 7 Then
vtx2 = vtx1
Else
vtx2 = ""
End If

End Sub
Public Sub Sauvegarde()

End Sub
