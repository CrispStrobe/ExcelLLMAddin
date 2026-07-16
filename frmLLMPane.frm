VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLLMPane 
   Caption         =   "LLM Assistant"
   ClientHeight    =   11040
   ClientLeft      =   100
   ClientTop       =   420
   ClientWidth     =   6520
   OleObjectBlob   =   "frmLLMPane.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmLLMPane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    modPane.Pane_Init Me
End Sub
Private Sub cboProvider_Change()
    modPane.Pane_ProviderChanged Me
End Sub
Private Sub btnLoadModels_Click()
    modPane.Pane_LoadModels Me
End Sub
Private Sub btnSave_Click()
    modPane.Pane_Save Me
End Sub
Private Sub btnTest_Click()
    modPane.Pane_Test Me
End Sub
Private Sub btnRunPrompt_Click()
    modPane.Pane_RunPrompt Me
End Sub
Private Sub btnRunAgent_Click()
    modPane.Pane_RunAgent Me
End Sub
