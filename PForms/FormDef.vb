'***********************************************************************
'Copyright 2005-2022 Vijay Sridhara

'Licensed under the Apache License, Version 2.0 (the "License");
'you may not use this file except in compliance with the License.
'You may obtain a copy of the License at

'http://www.apache.org/licenses/LICENSE-2.0

'Unless required by applicable law or agreed to in writing, software
'distributed under the License is distributed on an "AS IS" BASIS,
'WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'See the License for the specific language governing permissions and
'limitations under the License.
'***********************************************************************
Public Class FormDef
    Implements IElement

    Private _title As String
    Private _fields As New Dictionary(Of String, Field)

    Private _elementtype As PDElementType
    Private _hasallLabels As Boolean
    Private _datastore As String
    Private _VERSION As String
    Private _USERNAME As String
    Public Property AllFieldsLabels() As Boolean
        Get
            Return _hasallLabels
        End Get
        Set(ByVal value As Boolean)
            _hasallLabels = value
        End Set
    End Property
    Public Property User As String
        Get
            Return _USERNAME
        End Get
        Set(ByVal value As String)
            _USERNAME = value
        End Set
    End Property
    Public Property DataStore As String
        Get
            Return _datastore
        End Get
        Set(ByVal value As String)
            _datastore = value
        End Set
    End Property

    Public ReadOnly Property ElementType As PDElementType Implements IElement.ElementType
        Get
            Return PDElementType.FormDefinition
        End Get

    End Property
    Public Property Title As String
        Get
            Return _title
        End Get
        Set(ByVal value As String)
            _title = value
        End Set
    End Property

    Public ReadOnly Property Fields() As Dictionary(Of String, Field)
        Get
            Return _fields
        End Get
    End Property


    Public Property Ver As String
        Get
            Return _VERSION
        End Get
        Set(ByVal value As String)
            _VERSION = value
        End Set
    End Property






End Class
