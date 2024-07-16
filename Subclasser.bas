Attribute VB_Name = "Subclasser"
Option Explicit

Private Const WM_NCDESTROY As Long = &H82&
Private Const WM_UAHDESTROYWINDOW As Long = &H90& 'Undocumented.

Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" ( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As Long, _
    ByVal uIdSubclass As Long, _
    ByVal dwRefData As Long) As Long

Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" ( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As Long, _
    ByVal uIdSubclass As Long) As Long

Public Declare Function DefSubclassProc Lib "comctl32" Alias "#413" ( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Function SubclassMe( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As rDIconConfigForm, _
    Optional ByVal dwRefData As Long) As Boolean
    
    SubclassMe = SetWindowSubclass(hWnd, _
                                   AddressOf SubclassProxy, _
                                   ObjPtr(pfnSubclass), _
                                   dwRefData)
End Function

Public Function RemoveMe( _
    ByVal hWnd As Long, _
    ByVal pfnSubclass As rDIconConfigForm) As Boolean
    
    RemoveMe = RemoveWindowSubclass(hWnd, _
                                    AddressOf SubclassProxy, _
                                    ObjPtr(pfnSubclass))
End Function

Private Function SubclassProxy( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long, _
    ByVal uIdSubclass As rDIconConfigForm, _
    ByVal dwRefData As Long) As Long
   
    If uMsg = WM_NCDESTROY Or uMsg = WM_UAHDESTROYWINDOW Then
        'Just in case the client fails to clean up.
        RemoveMe hWnd, uIdSubclass
    Else
        SubclassProxy = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
    End If
End Function
