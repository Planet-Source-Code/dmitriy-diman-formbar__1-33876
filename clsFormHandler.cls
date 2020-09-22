VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CLASS MODULE THAT HANDLES ACTIVATE_EVENT OF A FORM ASSIGNED TO IT

Private WithEvents tmpForm As Form
Attribute tmpForm.VB_VarHelpID = -1
Private UC_Owner As FormBar

Private F_Stat As Boolean 'Active Or not

Public Sub setOwner(UC)
    Set UC_Owner = UC
End Sub

Public Sub AssignFormToClass(frm)
    Set tmpForm = frm
End Sub

Private Sub tmpForm_Activate()
    'Activate event
    F_Stat = True
    UC_Owner.Forms_Activate_Event tmpForm.hDC
End Sub


Public Function getHDC() As Long
    getHDC = tmpForm.hDC
End Function

Public Function getCaption() As String
    getCaption = tmpForm.Caption
End Function

Private Sub tmpForm_Deactivate()
    F_Stat = False
End Sub

Private Sub tmpForm_Unload(Cancel As Integer)
    DoEvents 'BECAUSE WHEN FORM IC DEACTIVTED AND WE TRY TO CLOSE IT TWO EVENTS FIRE AT THE SAME TIME
    UC_Owner.Forms_Unload_Event tmpForm.hDC
End Sub


Public Function getFState() As Boolean
    getFState = F_Stat
End Function

Public Function setActive()
    tmpForm.ZOrder 0
End Function

Public Function getFormState() As Integer
    getFormState = tmpForm.WindowState
End Function
Public Sub setFormState(state As Integer)
    tmpForm.WindowState = state
End Sub
