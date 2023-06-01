'/********************************************************************/
'/* $RCSfile: CIHash.vb,v $ */
'/* $Date: 2006/12/27 16:12:01 $ */
'/* $Revision: 1.4 $ */
'/* $Author: mike $ */
'/********************************************************************/
'  Make a case insensitve Hashtable
Option Explicit On
Option Strict On
Imports System
Imports System.Collections
Public Class CIHashtable
    Inherits Hashtable

    Dim m_aryKeysLC As ArrayList
    Dim m_aryKeys As ArrayList
    Dim m_bIsCaseSensitive As Boolean

    ' --------------------------------------------------------------------
    Public Sub New()
        Me.New(False)
    End Sub
    Public Sub New(ByVal nCapacity As Integer)
        Me.New(nCapacity, False)
    End Sub
    Public Sub New(ByVal nCapacity As Integer, ByVal IsCaseSensitive As Boolean)
        MyBase.New(nCapacity)
        m_aryKeysLC = New ArrayList
        m_aryKeys = New ArrayList
        m_bIsCaseSensitive = IsCaseSensitive
    End Sub
    Public Sub New(ByVal IsCaseSensitive As Boolean)
        MyBase.New()
        m_aryKeysLC = New ArrayList
        m_aryKeys = New ArrayList
        m_bIsCaseSensitive = IsCaseSensitive
    End Sub
    Public Sub New(ByVal IsCaseSensitive As Boolean, ByVal d As IDictionary)
        MyBase.New(d)
        m_aryKeysLC = New ArrayList
        m_aryKeys = New ArrayList
        m_bIsCaseSensitive = IsCaseSensitive
    End Sub
    Public Sub New(ByVal d As IDictionary)
        Me.New(False, d)
    End Sub
    Public Sub New(ByVal ht As CIHashtable)
        MyBase.New(ht)
        m_aryKeysLC = New ArrayList
        m_aryKeys = New ArrayList

        Dim de As IDictionaryEnumerator = MyBase.GetEnumerator
        While de.MoveNext
            If Me.m_bIsCaseSensitive = False AndAlso de.Key.GetType Is GetType(System.String) Then
                Dim strKey As String = CStr(de.Key)
                m_aryKeysLC.Add(strKey.ToLower)
            Else
                m_aryKeysLC.Add(de.Key)
            End If
            m_aryKeys.Add(de.Key)
        End While
        m_bIsCaseSensitive = ht.IsCaseSensitive
    End Sub
    Public Sub New(ByVal ht As Hashtable)
        MyBase.New(ht)
        m_aryKeysLC = New ArrayList
        m_aryKeys = New ArrayList
        For Each dict As DictionaryEntry In ht
            If dict.Key.GetType Is GetType(System.String) Then
                m_aryKeysLC.Add(CStr(dict.Key).ToLower)
            Else
                m_aryKeysLC.Add(dict.Key)
            End If
            m_aryKeys.Add(dict.Key)
        Next
        m_bIsCaseSensitive = False
    End Sub

    Public ReadOnly Property IsCaseSensitive() As Boolean
        Get
            Return m_bIsCaseSensitive
        End Get
    End Property
    Private Function GetKeyOfs(ByVal key As Object) As Integer
        Try
            If Me.m_aryKeysLC.Count = 0 Then
                Throw New Exception("Key not found in collection. Collection is empty")
            End If
            If (Me.m_bIsCaseSensitive = False) AndAlso (key.GetType Is GetType(System.String)) Then
                Dim strKeyLC As String = CStr(key).ToLower
                For nCnt As Integer = 0 To Me.m_aryKeysLC.Count - 1
                    If Me.m_aryKeysLC(nCnt).GetType Is GetType(System.String) Then
                        If strKeyLC = CStr(m_aryKeysLC(nCnt)) Then
                            Return nCnt
                        End If
                    End If
                Next
            Else
                Dim nCnt As Integer = 0

                Dim de As IDictionaryEnumerator = MyBase.GetEnumerator
                While de.MoveNext
                    If Me.m_aryKeys(nCnt) Is key Then
                        Return nCnt
                    End If
                    nCnt = nCnt + 1
                End While
            End If
        Catch ex As Exception
            Debug.Assert(False, ex.ToString)
        End Try
        Return -1
    End Function

    Default Public Overrides Property Item(ByVal key As Object) As Object
        Get
            If key Is Nothing Then
                Throw New Exception("Key is NULL. Key not not valid")
            End If
            If Me.m_aryKeysLC.Count = 0 Then
                Return Nothing
            End If
            Dim nCnt As Integer = GetKeyOfs(key)
            If nCnt >= 0 Then
                Return MyBase.Item(Me.m_aryKeys(nCnt))
            End If
            Return Nothing
        End Get
        Set(ByVal Value As Object)
            Static nItter As Integer
            nItter += 1
            If Me.m_aryKeysLC.Count = 0 Then
                Throw New Exception("Key not found in collection. Collection is empty")
            End If
            If key Is Nothing Then
                Throw New Exception("Key is NULL. Key not not valid")
            End If
            Try
                Dim nCnt As Integer = GetKeyOfs(key)
                MyBase.Item(Me.m_aryKeys(nCnt)) = Value
            Catch ex As Exception
                Debug.Assert(False, ex.ToString)
            End Try
            nItter -= 1
        End Set
    End Property

    Public Overloads Sub Remove(ByVal key As Object)
        Dim nCnt As Integer = GetKeyOfs(key)
        MyBase.Remove(m_aryKeys(nCnt))
        m_aryKeysLC.RemoveAt(nCnt)
        m_aryKeys.RemoveAt(nCnt)
    End Sub

    Public Shared Sub Swap(ByRef ht1 As CIHashtable, ByRef ht2 As CIHashtable)
        Dim htTmp As CIHashtable = New CIHashtable(ht1)
        ht1 = New CIHashtable(ht2)
        ht2 = New CIHashtable(htTmp)
    End Sub

    Public Overloads Shared Sub Overlay(ByRef htTarget As CIHashtable, ByRef htOverlay As CIHashtable)
        If htOverlay Is Nothing Then Return
        If htTarget Is Nothing Then
            htTarget = New CIHashtable(htOverlay)
            Return
        End If
        For Each d As DictionaryEntry In htOverlay
            If htTarget.ContainsKey(d.Key) Then
                htTarget(d.Key) = d.Value
            Else
                htTarget.Add(d.Key, d.Value)
            End If
        Next
    End Sub
    Public Overloads Sub Add(ByVal key As Object, ByVal value As Object)
        ' If IsCaseSensitive = False, and key is string, and key exists: it is updated
        If (Me.m_bIsCaseSensitive = False) AndAlso (key.GetType Is GetType(System.String)) Then
            key = CStr(key).Trim
            If CStr(key) = "" Then
                Throw New Exception("Invalid Key. Key cannot be null string or only whitespace")
            End If
            Dim strKeyLC As String = CStr(key).ToLower
            If Not m_aryKeysLC.Contains(strKeyLC) Then
                m_aryKeysLC.Add(strKeyLC)
                m_aryKeys.Add(key)
                MyBase.Add(key, value)
            Else
                Dim nCnt As Integer = GetKeyOfs(key)
                MyBase.Item(Me.m_aryKeys(nCnt)) = value
            End If
        Else
            MyBase.Add(key, value)
            m_aryKeysLC.Add(key)
            m_aryKeys.Add(key)
        End If
    End Sub
    Public Overloads Sub Clear()
        MyBase.Clear()
        m_aryKeysLC.Clear()
        m_aryKeys.Clear()
    End Sub

    Public Overloads Function Contains(ByVal key As Object) As Boolean
        If (m_bIsCaseSensitive = False) AndAlso (key.GetType Is GetType(System.String)) Then
            key = CStr(key).Trim
            If CStr(key) = "" Then
                Throw New Exception("Invalid Key. Key cannot be null string or only whitespace")
            End If
            Return m_aryKeysLC.Contains(CStr(key).ToLower)
        Else
            Return MyBase.Contains(key)
        End If
    End Function

    Public Overloads Function ContainsKey(ByVal key As Object) As Boolean
        If (m_bIsCaseSensitive = False) AndAlso (key.GetType Is GetType(System.String)) Then
            key = CStr(key).Trim
            If CStr(key) = "" Then
                Throw New Exception("Invalid Key. Key cannot be null string or only whitespace")
            End If
            Return m_aryKeysLC.Contains(CStr(key).ToLower)
        Else
            Return MyBase.Contains(key)
        End If
    End Function

End Class ' CIHash