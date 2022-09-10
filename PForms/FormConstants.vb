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
Public Module FormConstants
    Enum PDElementType
        FormDefinition
        FormField
        FormReferences
        FormReport
        FormMeta
    End Enum

    Enum PDFieldType
        User
        Reference
        List
        Text
        Multiline
        Largetext
        Check
        [Date]
        Radio
        File
        Picture
        Combo
        ROCombo
        Numeric
        Label
        Timestamp
        MonthYear
        Serial
        RichText
    End Enum
    Public Const FORM_VER = "0.2.4"
End Module
