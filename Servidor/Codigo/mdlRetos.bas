Attribute VB_Name = "mdlRetos"
Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = TIndex
    UserList(TIndex).flags.EstaDueleando = True
    Call WarpUserChar(TIndex, 205, 41, 55)
    UserList(TIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 205, 63, 59)
    Call SendData(ToAll, 0, 0, "||Retos 1 vs 1> " & UserList(TIndex).Name & " y " & UserList(UserIndex).Name & " van a competir en un Reto." & FONTTYPE_TALK)
End Sub
Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    Call WarpUserChar(UserIndex, 59, 52, 41)
    Call WarpUserChar(TIndex, 59, 52, 42)
    UserList(TIndex).flags.EsperandoDuelo = False
    UserList(TIndex).flags.Oponente = 0
    UserList(TIndex).flags.EstaDueleando = False
End Sub
Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos 1 vs 1> " & UserList(Ganador).Name & " venció a " & UserList(Perdedor).Name & " en un reto." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
End Sub
Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||Retos 1 vs 1> El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).Name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
End Sub
