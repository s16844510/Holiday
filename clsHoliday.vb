'================================================
' Copyright 2016 S16844510 All rights reserved.
'================================================

Public Class clsHoliday

#Region "Variable"

    Private m_Ical As String = String.Empty
    Private m_Cal As DataTable = Nothing

#End Region

#Region "Const"

    Private Const url As String = "https://calendar.google.com/calendar/ical/ja.japanese%23holiday%40group.v.calendar.google.com/public/basic.ics"

#End Region

#Region "### Constructer ###"
    Public Sub New()
        m_Cal = New DataTable
        m_Cal.Columns.Add("Year")
        m_Cal.Columns.Add("Month")
        m_Cal.Columns.Add("Day")
        m_Cal.Columns.Add("Summary")
    End Sub
#End Region
    Public Sub getData()
        getIcalData()
        formatCalData()
        outPutCalData()
    End Sub

#Region "### (getIcalData) ###"
    Private Sub getIcalData()

        Dim wc As New System.Net.WebClient()
        Try
            'set encoding UTF-8
            wc.Encoding = System.Text.Encoding.UTF8

            'download data as string
            m_Ical = wc.DownloadString(url)
            'len = m_Ical.Length

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            wc.Dispose()
        End Try
    End Sub
#End Region

#Region "### (formatCalData) ###"
    Private Sub formatCalData()

        Try
            For Each tmpGoup As String In Split(m_Ical, "END:VEVENT")
                Dim tmpRow As DataRow = m_Cal.NewRow()

                For Each tmpline As String In Split(tmpGoup, vbCrLf)
                    If tmpline.IndexOf("DTSTART") = 0 Then
                        Dim tmp As String
                        tmp = nthField(tmpline, ":", 2)
                        tmpRow("Year") = tmp.Substring(0, 4)
                        tmpRow("Month") = tmp.Substring(4, 2)
                        tmpRow("Day") = tmp.Substring(6, 2)
                    ElseIf tmpline.IndexOf("SUMMARY") = 0 Then
                        tmpRow("Summary") = nthField(tmpline, ":", 2)

                    End If
                Next
                m_Cal.Rows.Add(tmpRow)
            Next

            Dim view As DataView = New DataView(m_Cal)
            view.Sort = "Year,Month,Day"
            m_Cal = view.ToTable()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
#End Region

#Region "### (outPutCalData) ###"
    Private Sub outPutCalData()
        Dim outStream As New IO.StreamWriter("C:\Users\user\Desktop\out.csv")
        Dim tmp As String
        Try
            For Each dtRow As DataRow In m_Cal.Rows
                tmp = dtRow(0) + "," + dtRow(1) + "," + dtRow(2) + "," + dtRow(3)
                outStream.WriteLine(tmp)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            outStream.Dispose()
        End Try
    End Sub
#End Region

#Region "### (nthField) ###"
    Private Function nthField(str As String, sep As String, field As Integer) As String
        Dim line() As String = Split(str, sep)


        Return line(field - 1)
    End Function
#End Region
End Class
