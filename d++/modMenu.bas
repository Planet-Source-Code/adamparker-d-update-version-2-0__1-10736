Attribute VB_Name = "modMenu"
Option Explicit
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

    menu_handle = GetMenu(frm.hwnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

    With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With
    
    SetMenuItemInfo menu_handle, _
        item_numbers(UBound(item_numbers)), _
        True, menu_info
End Sub

