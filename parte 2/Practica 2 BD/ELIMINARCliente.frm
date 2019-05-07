VERSION 5.00
Begin VB.Form ELIMINARCliente 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ELIMINARCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()



If MsgBox("se eliminara el Usuario " & vCodigoCliente, vbInformation + vbYesNo, "AVISO") = vbYes Then

With rs_spEliminar

'PREGUNTAR A GASTY SOBRE EL METODO REQUERY Y EL POR QUE ME SALE OBJETO CERRADO
'        .Requery
'        .Find "[CÓDIGO CLIENTE] =' " & vCodigoCliente & " ' "
       ' If .EOF Then 'sino encuentra al cliente
        
        
        .Open " Execute sp_ElimarCliente '" & vCodigoCliente & "' ", cn, adOpenStatic, adLockOptimistic
        MsgBox "Se ELIMINO CLIENTE CORRECTAMENTE"
       
        
'        Else
'
'        MsgBox "NO EXISTE EL CLIENTE"
'        End If
        
End With


End If


End Sub
