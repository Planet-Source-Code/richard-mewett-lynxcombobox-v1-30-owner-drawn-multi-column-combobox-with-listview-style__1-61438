Attribute VB_Name = "mDemodata"
Option Explicit

Public Enum NameTypeEnum
    ntRandom = 0
    ntMale = 1
    ntFemale = 2
End Enum

#If False Then
    Private ntRandom, ntMale, ntFemale
#End If

'These are for generating the demo data
Private Const M_FORENAMES = "Andrew,Brian,David,Gavin,James,Jon,Matthew,Michael,Paul,Peter,Richard,Simon"
Private Const F_FORENAMES = "Amanda,Caroline,Charlotte,Jane,Karen,Katie,Kim,Lara,Lucy,Paula,Rachel,Sarah,Susan"
Private Const SURNAMES = "Anderson-Allen,Black,Bloggs,Brown,Cole,Davis,Lee,Ryan,Smith,Stephens,White"
   
Private mCalled As Boolean

Private sM() As String
Private sF() As String
Private sSurnames() As String
 
Public Function GetSurname() As String
    Initialise
    
    GetSurname = sSurnames(RandomInt(LBound(sSurnames), UBound(sSurnames)))
End Function


Public Function GetForename(Optional nType As NameTypeEnum) As String
    Initialise
    
    Select Case nType
        Case ntRandom
            If RandomInt(0, 1) = 0 Then
                GetForename = sM(RandomInt(LBound(sM), UBound(sM)))
            Else
                GetForename = sF(RandomInt(LBound(sF), UBound(sF)))
            End If
           
        Case ntMale
            GetForename = sM(RandomInt(LBound(sM), UBound(sM)))

        Case ntFemale
            GetForename = sF(RandomInt(LBound(sF), UBound(sF)))

    End Select
End Function



Public Function GetNameOfPerson(Optional nType As NameTypeEnum) As String
    Select Case nType
        Case ntRandom
            If RandomInt(0, 1) = 0 Then
                GetNameOfPerson = GetForename(ntMale) & " " & GetSurname()
            Else
                GetNameOfPerson = GetForename(ntFemale) & " " & GetSurname()
            End If
           
        Case ntMale
            GetNameOfPerson = GetForename(ntMale) & " " & GetSurname()

        Case ntFemale
            GetNameOfPerson = GetForename(ntFemale) & " " & GetSurname()

    End Select
End Function



Private Sub Initialise()
    If Not mCalled Then
        mCalled = True
        
        sM() = Split(M_FORENAMES, ",")
        sF() = Split(F_FORENAMES, ",")
        sSurnames() = Split(SURNAMES, ",")
    End If
End Sub


Public Function RandomInt(lowerbound As Long, upperbound As Long) As Long
    RandomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function



