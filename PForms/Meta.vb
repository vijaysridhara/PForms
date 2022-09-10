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
Public Class Meta
    Implements IElement
    Private _elementtype As PDElementType = PDElementType.FormMeta
    Private _name As String
    Private _ver As String
    Private _user As String
    Private _FDcOLL As New Dictionary(Of String, FormDef)
    Private _report As New Dictionary(Of String, Report)

    Public ReadOnly Property FormsVer As String
        Get
            Return FORM_VER
        End Get
    End Property
    Public Property Name As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property
    Public Property Version As String
        Get
            Return _ver
        End Get
        Set(ByVal value As String)
            _ver = value
        End Set
    End Property
    Public Property Username As String
        Get
            Return _user
        End Get
        Set(ByVal value As String)
            _user = value
        End Set
    End Property
    Public ReadOnly Property Forms() As Dictionary(Of String, FormDef)
        Get
            Return _FDcOLL
        End Get
    End Property
    Public ReadOnly Property GridReport As Dictionary(Of String, Report)
        Get
            Return _report
        End Get

    End Property

    Public ReadOnly Property ElementType As PDElementType Implements IElement.ElementType
        Get
            Return _elementtype
        End Get
    End Property
End Class
