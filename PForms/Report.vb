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
Public Class Report
    Implements IElement
    Private _elementtype As PDElementType
    Private _title As String
    Private _headers As String
    Private _inbuilt As Boolean
    Private _importDs As String
    Public Property ImportDataStore As String
        Get
            Return _importDs
        End Get
        Set(ByVal value As String)
            _importDs = value
        End Set
    End Property
    Public Property Inbuilt As Boolean
        Get
            Return _inbuilt
        End Get
        Set(ByVal value As Boolean)
            _inbuilt = value
        End Set
    End Property
    Public ReadOnly Property ElementType As PDElementType Implements IElement.ElementType
        Get
            Return PDElementType.FormReport
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
    Public Property SQL As String
        Get
            Return _headers
        End Get
        Set(ByVal value As String)
            _headers = value
        End Set
    End Property




End Class
