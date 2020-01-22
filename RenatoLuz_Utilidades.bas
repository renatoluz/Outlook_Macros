Attribute VB_Name = "RenatoLuz_Utilidades"
Sub A_Salvar_Backup()
file_name = ActiveWorkbook.name
file_path = ActiveWorkbook.Path


ActiveWorkbook.Save
ActiveWorkbook.SaveCopyAs "Y:\Demand & Portfolio Management\3 - Operation\2020\ControlesGovernanca\BACKUP\Import_WO_Control_" & Format(Now(), "YYYYMMDD_HHM") & ".xlsm"
End Sub







Sub SUBST_NOME_VENDORS()

Range("B10").Select
    Cells.Replace What:="7000064BRQ", Replacement:="BRQ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("B10").Select
    Cells.Replace What:="7000064MOOVEN", Replacement:="MOOVEN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("B10").Select
    Cells.Replace What:="7000000SENIORSOLUTION", Replacement:="SENIORSOLUTION", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7003718CPQI", Replacement:="CPQI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
        
    Range("B10").Select
    Cells.Replace What:="7000085ABACO - VAR3F", Replacement:="ABACO", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    
    Range("B10").Select
    Cells.Replace What:="7002865Altran", Replacement:="ALTRAN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


    Range("B10").Select
    Cells.Replace What:="7000000FinanceIT", Replacement:="FINANCE IT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


    Range("B10").Select
    Cells.Replace What:="10079024TresCon", Replacement:="TRES_CON", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False


    Range("B10").Select
    Cells.Replace What:="7000094Algartech", Replacement:="ALGARTECH", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    
    Range("B10").Select
    Cells.Replace What:="7000094Lan Designers", Replacement:="LAN_DESIGNERS", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    
    Range("B10").Select
    Cells.Replace What:="7000094Kiman", Replacement:="KIMAN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    
    Range("B10").Select
    Cells.Replace What:="7001000M2M", Replacement:="M2M", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Range("B10").Select
    Cells.Replace What:="7003718TresCon", Replacement:="TRES_CON", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Range("B10").Select
    Cells.Replace What:="7000000EBIX", Replacement:="EBIX", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("B10").Select
    Cells.Replace What:="10080350RSI", Replacement:="RSI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    Range("B10").Select
    Cells.Replace What:="7000002Beyond", Replacement:="BEYOND", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000005HDI", Replacement:="HDI", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000062Cedro", Replacement:="CEDROTECH", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
        
    Range("B10").Select
    Cells.Replace What:="7000064Resource", Replacement:="RESOURCE IT", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000069Accenture", Replacement:="ACCENTURE", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000078ALFAPEOPLE", Replacement:="ALFAPEOPLE", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000078Elumini", Replacement:="ELUMINI", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000080Realtask", Replacement:="REALTASK", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000082BTB Telecom", Replacement:="BTB TELECOM", _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000087Tivit", Replacement:="TIVIT", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7003515ATS", Replacement:="ATS", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7003718Advanta", Replacement:="ADVANTA", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7003718Cadmus", Replacement:="CADMUS", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("B10").Select
    Cells.Replace What:="7000064BRQ", Replacement:="BRQ", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

    Range("B10").Select
    Cells.Replace What:="7003718Solve", Replacement:="SOLVE", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Range("B10").Select
    Cells.Replace What:="10071534Mazzatech", Replacement:="MAZZATECH", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    Range("B10").Select
    Cells.Replace What:="7000094IK Solutions", Replacement:="IK Solutions", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    Range("B1").Select
    Cells.Replace What:="ABACO - VAR3F", Replacement:="ABACO", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    
    Range("B1").Select
    Cells.Replace What:="Brit", Replacement:="BRIT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    
       
    Range("B1").Select
    Cells.Replace What:="Automate", Replacement:="AUTOMATE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    
    
    Range("B1").Select
    Cells.Replace What:="Triple S", Replacement:="TRIPLE S", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
        

    Range("B1").Select
    Cells.Replace What:="Compasso", Replacement:="COMPASSO", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Tridea", Replacement:="TRIDEA", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Stefanini", Replacement:="STEFANINI", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Resource", Replacement:="RESOURCE IT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Red Hat", Replacement:="RED HAT", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Orion", Replacement:="ORION", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Mastery", Replacement:="MASTERY", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="IK Solutions", Replacement:="IK SOLUTIONS", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Himself Associates", Replacement:="HIMSELF ASSOCIATES", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="FastFin", Replacement:="FASTFIN", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Everis", Replacement:="EVERIS", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="DBACorp", Replacement:="DBACORP", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Concrete", Replacement:="CONCRETE", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False


    Range("B1").Select
    Cells.Replace What:="Compasso", Replacement:="COMPASSO", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
        
    
    
        
End Sub

