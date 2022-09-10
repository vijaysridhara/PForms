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
Public Class Field
    Implements IElement
    Private _elementtype As PDElementType

    Private CTL As Object
    Public Property Control As Object
        Get
            Return CTL
        End Get
        Set(ByVal value As Object)
            CTL = value
        End Set
    End Property
    Private _name As String
    Private _ftype As PDFieldType
    Private _datasource As String
    Public ReadOnly Property ElementType As PDElementType Implements IElement.ElementType
        Get
            Return PDElementType.FormField
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
    Public Property FieldType As PDFieldType
        Get
            Return _ftype
        End Get
        Set(ByVal value As PDFieldType)
            _ftype = value
        End Set
    End Property
    Public Property DataSource As String
        Get
            Return _datasource
        End Get
        Set(ByVal value As String)
            _datasource = value
        End Set
    End Property



End Class
