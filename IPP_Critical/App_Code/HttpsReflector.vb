Imports Microsoft.VisualBasic
Imports System
Imports System.Web.Services.Description

Public Class HttpsReflector
    Inherits SoapExtensionReflector


    Public Overrides Sub ReflectMethod()
        ' No-OP
    End Sub

    Public Overrides Sub ReflectDescription()

        Dim description As ServiceDescription = ReflectionContext.ServiceDescription
        Dim service As Service
        Dim port As Port
        Dim extension As ServiceDescriptionFormatExtension


        'For Each service In description.Services

        '    For Each port In service.Ports


        '        For Each extension In port.Extensions

        '            Dim binding As SoapAddressBinding = extension as SoapAddressBinding

        '        Next

        '    Next

        'Next

    End Sub


End Class
