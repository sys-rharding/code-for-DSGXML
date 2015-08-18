Imports Microsoft.Office.Interop.Excel

Module Module1

    Sub formatSheet(ByRef excelapp)
        'Add Grouping support to datasheet generator utility
        'Written by Robert Harding, Harding Consulting Group.
        'Date April 8, 2013

        'Variables defined
        Dim objectfirstPos(7) As Integer
        Dim objectname(7) As String
        Dim previousvalue As String
        Dim x, y, pos, rb, dd, cb, ll, btn, ti, st, initpos, finalpos, columncount, wt As Integer
        Dim objectValue As String
        Dim workingsheet As Object
        Dim foo As Microsoft.Office.Interop.Excel.Application


        foo = excelapp

        For z = 1 To foo.Worksheets.Count

            pos = 0
            rb = 0
            cb = 0
            ll = 0
            ti = 0
            st = 0
            btn = 0
            dd = 0
            wt = 0
            columncount = 0


            workingsheet = foo.Worksheets(z)

            ' Columns = Nothing
            '  workingsheet = Sheet4

            workingsheet.Select()
            columncount = workingsheet.UsedRange.Columns.Count
            'x is counter that contains the column number of the cell
            x = 1

            previousvalue = ""
            While x <= columncount
                workingsheet.Range("A4").Offset(0, x - 1).Select()
                objectValue = CStr(workingsheet.Range("A4").Offset(0, x - 1).Value)



                If objectValue <> previousvalue Then

                    pos = pos + 1
                    objectfirstPos(pos - 1) = x
                    objectname(pos - 1) = objectValue

                    Debug.Print(objectValue)
                    Debug.Print(x)

                End If


                'These variables gets the total number of  objects per page
                If objectValue = "Radio Button" Then
                    rb = rb + 1
                End If
                If objectValue = "Drop-Down" Then
                    dd = dd + 1
                End If

                If objectValue = "Check-Box" Then
                    cb = cb + 1
                End If

                If objectValue = "Link" Then
                    ll = ll + 1
                End If

                If objectValue = "Button" Then
                    btn = btn + 1
                End If

                If objectValue = "Text Input" Then
                    ti = ti + 1
                End If

                If objectValue = "Screen Text" Then
                    st = st + 1
                End If

                If objectValue = "WebTable" Then
                    wt = wt + 1
                End If

                previousvalue = objectValue
                x = x + 1
            End While

            Debug.Print("")
            Debug.Print("")
            Debug.Print(st & " Screen Text")
            Debug.Print(rb & " Radio Button")
            Debug.Print(dd & " Drop-Down")
            Debug.Print(cb & " Check-Box")
            Debug.Print(ll & " Link")
            Debug.Print(btn & "Button")
            Debug.Print(ti & "Text Input")




            For y = 0 To 6

                Select Case objectname(y)

                    Case "Radio Button"
                        initpos = objectfirstPos(y)
                        finalpos = rb

                        If ((finalpos + initpos) - initpos) > 3 Then
                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If

                    Case "Drop-Down"
                        initpos = objectfirstPos(y)
                        finalpos = dd

                        If ((finalpos + initpos) - initpos) > 3 Then
                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If


                    Case "Check-Box"
                        initpos = objectfirstPos(y)
                        finalpos = cb

                        If ((finalpos + initpos) - initpos) > 3 Then
                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If


                    Case "Link"
                        initpos = objectfirstPos(y)
                        finalpos = ll

                        If (finalpos + initpos) > 3 Then
                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If


                    Case "Button"
                        initpos = objectfirstPos(y)
                        finalpos = btn

                        If ((finalpos + initpos) - initpos) > 3 Then
                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If



                    Case "Text Input"
                        initpos = objectfirstPos(y)
                        finalpos = ti

                        If ((finalpos + initpos) - initpos) > 3 Then

                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If


                    Case "WebTable"
                        initpos = objectfirstPos(y)
                        finalpos = wt

                        If ((finalpos + initpos) - initpos) > 3 Then

                            With workingsheet
                                .Range(.Columns(initpos), .Columns((finalpos + initpos) - 2)).Group()
                            End With
                        End If

                    Case "Screen Text"
                        initpos = objectfirstPos(y)
                        finalpos = st

                        If (finalpos - initpos) > 3 Then

                            With workingsheet
                                .Range(.Columns(initpos), .Columns(finalpos - 1)).Group()
                            End With
                        End If

                End Select

            Next y
        Next z
    End Sub



End Module


