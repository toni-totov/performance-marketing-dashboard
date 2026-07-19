"""frmExchangeRate -- the Monthly Media Plan exchange-rate popup (PLAN.md
§9.4). A single numeric TextBox, validated as a positive number, that
writes to the workbook-scoped MonthlyExchangeRate named cell on OK;
Cancel discards.

Built via sheets/forms/_builder.py's build_form_from_spec() at injection
time -- see that module's docstring and PLAN.md §18 for why this isn't a
static .frm file.

Calling convention from Monthly Media Plan's Worksheet_BeforeDoubleClick
handler:
    frmExchangeRate.ShowRate
ShowRate/PrepareRate are split the same way as frmNotes' ShowForCell/
PrepareForCell (Phase 1) so the read side is testable without ever
triggering the blocking modal Me.Show. The write side needs its own
split beyond frmNotes' plain CommitAndClose, since -- unlike a Notes
field, which accepts any text -- an exchange rate must validate as a
positive number: TryApplyRate() is the testable, MsgBox-free core (valid
input writes the rate and returns True; invalid input leaves the named
cell untouched and returns False); btnOK_Click is the thin
MsgBox-showing production wrapper around it (Unloads on success, warns
and stays open on failure) -- same Public-for-modal-bypass reasoning as
Tasks' CascadeDelete/WouldCreateCycle split (PLAN.md §19 Phase 3).
"""

SPEC = {
    "name": "frmExchangeRate",
    "controls": [
        {
            "type": "Forms.Label.1",
            "name": "lblPrompt",
            "properties": {
                "Caption": "1 EUR =",
                "Left": 10,
                "Top": 14,
                "Width": 50,
                "Height": 18,
            },
        },
        {
            "type": "Forms.TextBox.1",
            "name": "txtRate",
            "properties": {
                "Left": 65,
                "Top": 10,
                "Width": 70,
                "Height": 20,
            },
        },
        {
            "type": "Forms.Label.1",
            "name": "lblUnit",
            "properties": {
                "Caption": "USD",
                "Left": 140,
                "Top": 14,
                "Width": 40,
                "Height": 18,
            },
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnOK",
            "properties": {"Caption": "OK", "Left": 65, "Top": 45, "Width": 55, "Height": 24},
        },
        {
            "type": "Forms.CommandButton.1",
            "name": "btnCancel",
            "properties": {
                "Caption": "Cancel",
                "Left": 125,
                "Top": 45,
                "Width": 60,
                "Height": 24,
                "Cancel": True,
            },
        },
    ],
    "code": """
Private Sub UserForm_Initialize()
    Me.Caption = "Set Exchange Rate"
    Me.Width = 220
    Me.Height = 115
End Sub

Public Sub ShowRate()
    PrepareRate
    Me.Show
End Sub

' Format(), not TEXT() -- see CLAUDE.md's TEXT()-locale lesson. VBA's
' Format() function's "0.00" format-string tokens are locale-independent
' (unlike Excel's TEXT() worksheet function, which parses the format
' string itself against the current locale), and pairs correctly with
' CDbl()'s own locale-aware parsing in TryApplyRate below.
Public Sub PrepareRate()
    Dim currentRate As Double
    currentRate = ThisWorkbook.Names("MonthlyExchangeRate").RefersToRange.Value
    ' 4 decimals, not 2 (user-requested, PLAN.md §19 Phase 5) -- also
    ' fixes a latent precision-loss bug: reopening this popup and
    ' clicking OK without retyping would previously overwrite a
    ' 4-decimal rate with its 2-decimal-rounded display text.
    txtRate.Text = Format(currentRate, "0.0000")
End Sub

' Returns True and writes MonthlyExchangeRate if txtRate holds a positive
' number; returns False (named cell untouched) otherwise. No MsgBox here
' -- see module docstring for why this is split from btnOK_Click.
Public Function TryApplyRate() As Boolean
    Dim txt As String
    txt = Trim(txtRate.Text)

    If Not IsNumeric(txt) Then
        TryApplyRate = False
        Exit Function
    End If

    Dim val As Double
    val = CDbl(txt)
    If val <= 0 Then
        TryApplyRate = False
        Exit Function
    End If

    ThisWorkbook.Names("MonthlyExchangeRate").RefersToRange.Value = val
    TryApplyRate = True
End Function

Private Sub btnOK_Click()
    If TryApplyRate() Then
        Unload Me
    Else
        MsgBox "Please enter a positive number.", vbExclamation, "Invalid Exchange Rate"
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
""",
}
