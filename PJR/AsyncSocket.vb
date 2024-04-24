'ShadowMud.NET - Open Source Mud Server Framework
'
'Copyright (C) 2001-2004 
'   Tim Davis (darkmercenary@earthlink.net)

'This program is free software; you can redistribute it and/or modify it under
' the terms of the GNU General Public License as published by the Free Software
'Foundation; either version 2 of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful, but WITHOUT ANY
'WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A 
'PARTICULAR PURPOSE. See the GNU General Public License for more details.

'You should have received a copy of the GNU General Public License along with this
'program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, 
'Suite 330, Boston, MA 02111-1307 USA

Imports System
Imports System.Text
Imports System.Net
Imports System.Net.Sockets

Namespace ShadowMud.Sockets

#Region "StateObject"

    Public Class StateObject

        Public WorkSocket As Socket = Nothing

        Public BufferSize As Integer = 32767

        Public Buffer(32767) As Byte
        'Public BufferSize As Integer = 8192

        'Public Buffer(8192) As Byte


        Public StrBuilder As New StringBuilder

    End Class

#End Region

#Region "AsyncServer"

    Public Class AsyncServer

#Region "Variables"

        Private m_SocketPort As Integer

        Public Event ConnectionAccept(ByVal tmp_Socket As AsyncSocket)

        Private m_Close As Boolean

        Private utility As New Utility

#End Region

#Region "Methods"

        Public Sub New(ByVal tmp_Port As Integer)

            m_SocketPort = tmp_Port

        End Sub

        Public Sub Start()
            utility.TraceLogTxt("AsyncServer-Start")

            Dim strHostName As String
            strHostName = System.Net.Dns.GetHostName()

            Dim ListenIP As IPAddress = IPAddress.Any
            Dim ListenPort As Integer = m_SocketPort
            Dim ListenEP As IPEndPoint = New IPEndPoint(ListenIP, ListenPort)

            If m_Close = True Then m_Close = False : Exit Sub

            Dim obj_Socket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            Try

                obj_Socket.Bind(ListenEP)
                obj_Socket.Listen(100)
                obj_Socket.BeginAccept(New AsyncCallback(AddressOf onIncomingConnection), obj_Socket)

            Catch ex As SocketException
                utility.TraceLogTxt("AsyncServer-Start " & ex.Message & vbCrLf & ex.StackTrace)
            End Try

        End Sub

        Public Sub Close()

            m_Close = True

        End Sub

#End Region

#Region "Private Subs/Functions"

        Private Sub onIncomingConnection(ByVal ar As IAsyncResult)

            Try
                Dim obj_Socket As Socket = CType(ar.AsyncState, Socket)
                Dim obj_Connected As Socket = obj_Socket.EndAccept(ar)

                'Server is Shutdown
                If m_Close = True Then

                    'Shutdown the New Socket
                    obj_Connected.Shutdown(SocketShutdown.Both)
                    obj_Connected.Close()

                Else

                    'Get GUID for new socket
                    Dim tmp_GUID As String = GetGUID()

                    'Raise Event
                    RaiseEvent ConnectionAccept(New AsyncSocket(obj_Connected, tmp_GUID))

                End If

                obj_Socket.BeginAccept(New AsyncCallback(AddressOf onIncomingConnection), obj_Socket)
            Catch ex As SocketException
                utility.TraceLogTxt("AsyncServer-onIncomingConnection " & ex.Message & vbCrLf & ex.StackTrace)
            End Try
            

        End Sub

        Private Function GetGUID() As String

            Return System.Guid.NewGuid.ToString

        End Function

#End Region

    End Class

#End Region

#Region "AsyncSocket"

    Public Class AsyncSocket

#Region "Variables"

        Private m_SocketID As String

        Private m_IpClientID As String

        Private m_tmpSocket As Socket

        Private m_recBuffer As String

        Public Event socketDisconnected(ByVal SocketID As String, ByVal IpAddress As String)

        Public Event socketDataArrival(ByVal SocketID As String, ByVal SocketData As String, ByVal IpAddress As String)

        Public Event socketConnected(ByVal SocketID As String)

        Private utility As New Utility

#End Region

#Region "Methods"

        Public Sub New(ByVal tmp_Socket As Socket, ByVal tmp_SocketID As String)
            Try
                m_SocketID = tmp_SocketID
                m_tmpSocket = tmp_Socket

                m_IpClientID = tmp_Socket.RemoteEndPoint.ToString

                Dim obj_Socket As Socket = tmp_Socket
                Dim obj_SocketState As New StateObject

                obj_SocketState.WorkSocket = obj_Socket
                obj_Socket.BeginReceive(obj_SocketState.Buffer, 0, obj_SocketState.BufferSize, 0, New AsyncCallback(AddressOf onDataArrival), obj_SocketState)

            Catch ex As SocketException
                utility.TraceLogTxt("AsyncSocket-New" & ex.Message & vbCrLf & ex.StackTrace)
            End Try
           
        End Sub

        Public Sub New()
            m_tmpSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        End Sub

        Public Sub Send(ByVal tmp_Data As String)

            Try

                Dim obj_StateObject As New StateObject
                obj_StateObject.WorkSocket = m_tmpSocket

                Dim Buffer As Byte() = Encoding.Default.GetBytes(tmp_Data)
                m_tmpSocket.BeginSend(Buffer, 0, Buffer.Length, 0, New AsyncCallback(AddressOf onSendComplete), obj_StateObject)

            Catch ex As SocketException
                utility.TraceLogTxt("AsyncSocket-Send" & ex.Message & vbCrLf & ex.StackTrace)
            End Try

        End Sub

        Public Sub Close()

            m_tmpSocket.Shutdown(SocketShutdown.Both)
            m_tmpSocket.Close()

        End Sub

        Public Sub Connect(ByVal hostIP As String, ByVal hostPort As Integer)

            Dim hostEndPoint As New IPEndPoint(IPAddress.Parse(hostIP), hostPort)

            'Dim obj_Socket As New Socket(hostEndPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp)

            'm_tmpSocket = New Socket(hostEndPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp)

            Dim obj_Socket As Socket = m_tmpSocket

            Try
                obj_Socket.BeginConnect(hostEndPoint, New AsyncCallback(AddressOf onConnectionComplete), obj_Socket)
            Catch ex As SocketException
                utility.TraceLogTxt("AsyncSocket-Connect" & ex.Message & vbCrLf & ex.StackTrace)
            End Try


        End Sub

#End Region

#Region "Private Subs/Functions"

        Private Sub onDataArrival(ByVal ar As IAsyncResult)
            Dim ErrCode As Net.Sockets.SocketError
            Try

                Dim obj_SocketState As StateObject = CType(ar.AsyncState, StateObject)
                Dim obj_Socket As Socket = obj_SocketState.WorkSocket
                Dim sck_Data As String

                Dim BytesRead As Integer = obj_Socket.EndReceive(ar, ErrCode)
                If ErrCode <> SocketError.Success Then
                    BytesRead = 0
                End If

                If BytesRead > 0 Then
                    sck_Data = Encoding.Default.GetString(obj_SocketState.Buffer, 0, BytesRead)
                    RaiseEvent socketDataArrival(m_SocketID, sck_Data, m_IpClientID)
                End If

                'Start recieving again
                obj_Socket.BeginReceive(obj_SocketState.Buffer, 0, obj_SocketState.BufferSize, 0, New AsyncCallback(AddressOf onDataArrival), obj_SocketState)

            Catch e As SocketException
                utility.TraceLogTxt("AsyncSocket-onDataArrival" & e.Message & vbCrLf & e.StackTrace & vbCrLf & ErrCode.ToString)
                RaiseEvent socketDisconnected(m_SocketID, m_IpClientID)

            End Try

        End Sub

        Private Sub onSendComplete(ByVal ar As IAsyncResult)

            Dim obj_SocketState As StateObject = CType(ar.AsyncState, StateObject)
            Dim obj_Socket As Socket = obj_SocketState.WorkSocket

        End Sub

        Private Sub onConnectionComplete(ByVal ar As IAsyncResult)

            m_tmpSocket = CType(ar.AsyncState, Socket)
            m_tmpSocket.EndConnect(ar)

            RaiseEvent socketConnected("null")

            Dim obj_Socket As Socket = m_tmpSocket
            Dim obj_SocketState As New StateObject

            obj_SocketState.WorkSocket = obj_Socket
            obj_Socket.BeginReceive(obj_SocketState.Buffer, 0, obj_SocketState.BufferSize, 0, New AsyncCallback(AddressOf onDataArrival), obj_SocketState)

        End Sub

#End Region

#Region "Properties"

        Public ReadOnly Property SocketID() As String

            Get

                SocketID = m_SocketID

            End Get

        End Property

        Public ReadOnly Property IpClientID() As String

            Get

                IpClientID = m_IpClientID

            End Get

        End Property

#End Region

    End Class

#End Region

#Region "AsyncSocketController"

    Public Class AsyncSocketController

#Region "Variables"

        Private m_SocketCol As New SortedList

        Private WithEvents m_ServerSocket As AsyncServer

        Public Event onConnectionAccept(ByVal SocketID As String, ByVal IpAddress As String)

        Public Event onSocketDisconnected(ByVal SocketID As String, ByVal IpAddress As String)

        Public Event onDataArrival(ByVal SocketID As String, ByVal SocketData As String, ByVal IpAddress As String)

#End Region

#Region "Methods"

        Public Sub New(ByVal tmp_Port As Integer)

            m_ServerSocket = New AsyncServer(tmp_Port)

        End Sub

        Public Sub Start()

            m_ServerSocket.Start()

        End Sub

        Public Sub StopServer()

            m_ServerSocket.Close()

        End Sub

        Public Sub Send(ByVal tmp_SocketID As String, ByVal tmp_Data As String, Optional ByVal tmp_Return As Boolean = True)

            If tmp_Return = True Then

                CType(m_SocketCol.Item(tmp_SocketID), AsyncSocket).Send(tmp_Data & vbCrLf)

            Else

                CType(m_SocketCol.Item(tmp_SocketID), AsyncSocket).Send(tmp_Data)

            End If

        End Sub

        Public Sub Close(ByVal tmp_SocketID As String)

            CType(m_SocketCol.Item(tmp_SocketID), AsyncSocket).Close()

        End Sub

        Public Sub Add(ByVal tmp_Socket As AsyncSocket)

            m_SocketCol.Add(tmp_Socket.SocketID, tmp_Socket)

            AddHandler tmp_Socket.socketDisconnected, AddressOf SocketDisconnected
            AddHandler tmp_Socket.socketDataArrival, AddressOf SocketDataArrival

        End Sub

        Public ReadOnly Property Item(ByVal Key As String) As AsyncSocket
            Get
                Return m_SocketCol(Key)
            End Get
        End Property

        Public ReadOnly Property ItembyIndex(ByVal Key As Integer) As AsyncSocket
            Get
                Return m_SocketCol(m_SocketCol.GetKey(Key))
            End Get
        End Property

        Public Function Exists(ByVal SocketID As String) As Boolean

            Dim tmpSocket As ShadowMud.Sockets.AsyncSocket = Item(SocketID)

            If tmpSocket Is Nothing Then

                Return False

            Else

                Return True

            End If

        End Function

        Public ReadOnly Property Count()
            Get
                Return m_SocketCol.Count
            End Get
        End Property

#End Region

#Region "Private Subs/Functions"

        Private Sub m_ServerSocket_ConnectionAccept(ByVal tmp_Socket As AsyncSocket) Handles m_ServerSocket.ConnectionAccept

            Add(tmp_Socket)

            RaiseEvent onConnectionAccept(tmp_Socket.SocketID, tmp_Socket.IpClientID)

        End Sub

        Private Sub SocketDisconnected(ByVal SocketID As String, ByVal IpAddress As String)

            m_SocketCol.Remove(SocketID)

            RaiseEvent onSocketDisconnected(SocketID, IpAddress)

        End Sub

        Private Sub SocketDataArrival(ByVal SocketID As String, ByVal SocketData As String, ByVal IpAddress As String)

            RaiseEvent onDataArrival(SocketID, SocketData, IpAddress)

        End Sub

#End Region

    End Class

#End Region

End Namespace
