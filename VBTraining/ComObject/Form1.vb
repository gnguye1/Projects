Imports Microsoft.Office.Interop

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '必要な変数は Try　の外で宣言する
        Dim xlApplication As Excel.Application

        'COM　オブジェクトの解放を保証するために　Try ~ Finally を使用する
        Try
            xlApplication = New Excel.Application()

            '警告メッセージなどを表示しないようにする
            xlApplication.DisplayAlerts = False

            Dim xlBooks As Excel.Workbooks = xlApplication.Workbooks

            Try
                Dim xlBook As Excel.Workbook = xlBooks.Add()

                Try
                    Dim xlSheets As Excel.Sheets = xlBook.Worksheets

                    Try
                        Dim xlSheet As Excel.Worksheet = DirectCast(xlSheets(1), Excel.Worksheet)

                        Try
                            Dim xlCells As Excel.Range = xlSheet.Cells

                            Try
                                Dim xlRange As Excel.Range = DirectCast(xlCells(6, 4), Excel.Range)

                                Try
                                    'Microsoft Excel を表示する
                                    xlApplication.Visible = True

                                    '1000 ミリ秒(1秒) 待機する
                                    System.Threading.Thread.Sleep(1000)

                                    'Row=6, Column=4 の位置に文字をセットする
                                    xlRange.Value2 = "あと　1　秒で終了します。"

                                    '1000 ミリ秒(1秒) 待機する
                                    System.Threading.Thread.Sleep(1000)

                                    xlBook.SaveAs("C:\Users\CUCDUONG\Documents\Visual Studio 2013\Projects\VBTraining\ComObject\test_com.xlsx")
                                Finally
                                    If Not xlRange Is Nothing Then
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange)
                                    End If
                                End Try
                            Finally
                                If Not xlCells Is Nothing Then
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCells)
                                End If
                            End Try
                        Finally
                            If Not xlSheet Is Nothing Then
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet)
                            End If
                        End Try

                    Finally
                        If Not xlSheets Is Nothing Then
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets)
                        End If
                    End Try
                Finally
                    If Not xlBook Is Nothing Then
                        Try
                            xlBook.Close()
                        Finally
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook)
                        End Try
                    End If

                End Try
            Finally
                If Not xlBooks Is Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks)
                End If
            End Try

        Finally
            If Not xlApplication Is Nothing Then
                Try
                    xlApplication.Quit()

                Finally
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApplication)
                End Try
            End If
        End Try
    End Sub
End Class
