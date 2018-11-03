Public Class Form1

    ' Declarations

    Private Property FormatCounter As Object


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Turn off any warning message
        Label4.Visible = False
        ' Check the Clipboard DataType
        If My.Computer.Clipboard.ContainsText() Then
            ' Clipboard has Text data :
            ' - so put in the TextBox and make it visible and hide the Image Button
            TextBox1.Text = (My.Computer.Clipboard.GetText())
            TextBox1.Visible = True
            Button11.Visible = False
        ElseIf My.Computer.Clipboard.ContainsImage() Then
            ' Clipboard has image data : 
            ' - so put the image on a Button background and make it visible and the Textbox invisible
            Button11.BackgroundImage = (My.Computer.Clipboard.GetImage())
            Button11.Visible = True
            TextBox1.Visible = False
            ' now we have an image so enable the buttons to switch between text and image
            Button16.Visible = True
            Button17.Visible = True
        Else
            ' Clipboard has neither Text nor Image
            Label4.Text = " ! Clipboard contents too complex ! "
            Label4.Visible = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' Turn off any warning message
        Label4.Visible = False
        ' Check what type of data is in the box - based on visibility of controls (sneaky) !
        If TextBox1.Visible = True Then
            ' Is Text so copy to clipboard provided it's not empty
            If TextBox1.Text <> "" Then
                My.Computer.Clipboard.SetText(TextBox1.Text)
            Else
                Label4.Text = "! Box 1 is empty - Clipboard Not Changed !"
                Label4.Visible = True
            End If
        Else
            ' Must be an Image so copy to the clipboard
            My.Computer.Clipboard.SetImage(Button11.BackgroundImage)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check the Clipboard DataType
        If My.Computer.Clipboard.ContainsText() Then
            ' Clipboard has Text data :
            ' - so put in the TextBox and make it visible and hide the Image Button
            TextBox2.Text = (My.Computer.Clipboard.GetText())
            TextBox2.Visible = True
            Button12.Visible = False
        ElseIf My.Computer.Clipboard.ContainsImage() Then
            ' Clipboard has image data : 
            ' - so put the image on a Button background and make it visible and the Textbox invisible
            Button12.BackgroundImage = (My.Computer.Clipboard.GetImage())
            Button12.Visible = True
            TextBox2.Visible = False
            ' now we have an image so enable the buttons to switch between text and image
            Button18.Visible = True
            Button19.Visible = True
        Else
            ' Clipboard has neither Text nor Image
            Label4.Text = " ! Clipboard contents too complex ! "
            Label4.Visible = True
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check what type of data is in the box - based on visibility of controls (sneaky) !
        If TextBox2.Visible = True Then
            ' Is Text so copy to clipboard provided it's not empty
            If TextBox2.Text <> "" Then
                My.Computer.Clipboard.SetText(TextBox2.Text)
            Else
                Label4.Text = "! Box 2 is empty - Clipboard Not Changed !"
                Label4.Visible = True
            End If
        Else
            ' Must be an Image so copy to the clipboard
            My.Computer.Clipboard.SetImage(Button12.BackgroundImage)
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check the Clipboard DataType
        If My.Computer.Clipboard.ContainsText() Then
            ' Clipboard has Text data :
            ' - so put in the TextBox and make it visible and hide the Image Button
            TextBox3.Text = (My.Computer.Clipboard.GetText())
            TextBox3.Visible = True
            Button13.Visible = False
        ElseIf My.Computer.Clipboard.ContainsImage() Then
            ' Clipboard has image data : 
            ' - so put the image on a Button background and make it visible and the Textbox invisible
            Button13.BackgroundImage = (My.Computer.Clipboard.GetImage())
            Button13.Visible = True
            TextBox3.Visible = False
            ' now we have an image so enable the buttons to switch between text and image
            Button20.Visible = True
            Button21.Visible = True
        Else
            ' Clipboard has neither Text nor Image
            Label4.Text = " ! Clipboard contents too complex ! "
            Label4.Visible = True
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check what type of data is in the box - based on visibility of controls (sneaky) !
        If TextBox3.Visible = True Then
            ' Is Text so copy to clipboard provided it's not empty
            If TextBox3.Text <> "" Then
                My.Computer.Clipboard.SetText(TextBox3.Text)
            Else
                Label4.Text = "! Box 3 is empty - Clipboard Not Changed !"
                Label4.Visible = True
            End If
        Else
            ' Must be an Image so copy to the clipboard
            My.Computer.Clipboard.SetImage(Button13.BackgroundImage)
        End If

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check the Clipboard DataType
        If My.Computer.Clipboard.ContainsText() Then
            ' Clipboard has Text data :
            ' - so put in the TextBox and make it visible and hide the Image Button
            TextBox4.Text = (My.Computer.Clipboard.GetText())
            TextBox4.Visible = True
            Button14.Visible = False
        ElseIf My.Computer.Clipboard.ContainsImage() Then
            ' Clipboard has image data : 
            ' - so put the image on a Button background and make it visible and the Textbox invisible
            Button14.BackgroundImage = (My.Computer.Clipboard.GetImage())
            Button14.Visible = True
            TextBox4.Visible = False
            ' now we have an image so enable the buttons to switch between text and image
            Button22.Visible = True
            Button23.Visible = True
        Else
            ' Clipboard has neither Text nor Image
            Label4.Text = " ! Clipboard contents too complex ! "
            Label4.Visible = True
        End If

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check what type of data is in the box - based on visibility of controls (sneaky) !
        If TextBox4.Visible = True Then
            ' Is Text so copy to clipboard provided it's not empty
            If TextBox4.Text <> "" Then
                My.Computer.Clipboard.SetText(TextBox4.Text)
            Else
                Label4.Text = "! Box 4 is empty - Clipboard Not Changed !"
                Label4.Visible = True
            End If
        Else
            ' Must be an Image so copy to the clipboard
            My.Computer.Clipboard.SetImage(Button14.BackgroundImage)
        End If

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check the Clipboard DataType
        If My.Computer.Clipboard.ContainsText() Then
            ' Clipboard has Text data :
            ' - so put in the TextBox and make it visible and hide the Image Button
            TextBox5.Text = (My.Computer.Clipboard.GetText())
            TextBox5.Visible = True
            Button15.Visible = False
        ElseIf My.Computer.Clipboard.ContainsImage() Then
            ' Clipboard has image data : 
            ' - so put the image on a Button background and make it visible and the Textbox invisible
            Button15.BackgroundImage = (My.Computer.Clipboard.GetImage())
            Button15.Visible = True
            TextBox5.Visible = False
            ' now we have an image so enable the buttons to switch between text and image
            Button24.Visible = True
            Button25.Visible = True
        Else
            ' Clipboard has neither Text nor Image
            Label4.Text = " ! Clipboard contents too complex ! "
            Label4.Visible = True
        End If

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click

        ' Turn off any warning message
        Label4.Visible = False
        ' Check what type of data is in the box - based on visibility of controls (sneaky) !
        If TextBox5.Visible = True Then
            ' Is Text so copy to clipboard provided it's not empty
            If TextBox5.Text <> "" Then
                My.Computer.Clipboard.SetText(TextBox5.Text)
            Else
                Label4.Text = "! Box 5 is empty - Clipboard Not Changed !"
                Label4.Visible = True
            End If
        Else
            ' Must be an Image so copy to the clipboard
            My.Computer.Clipboard.SetImage(Button15.BackgroundImage)
        End If

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Dim Response As Integer

        Response = MsgBox("Exit Multi-Clip - are you sure ?", 4, "Multi-Clip Exit")
        ' Yes = 6, No = 7 

        If Response = 6 Then
            'MsgBox("Exiting")
            End
        End If

        MsgBox("Issue : form closes despite 'No' answer - sorry")
        'MsgBox("Not exiting")

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TextBox1.Text = "1"
        TextBox2.Text = "2"
        TextBox3.Text = "3"
        TextBox4.Text = "4"
        TextBox5.Text = "5"
        'TextBox4.Text = DateTime.Today
        'TextBox5.Text = DateTime.Now


        ' MsgBox("Current issue: cannot copy empty text into clipboard - fixed")

    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        ' make Image visible (and text invisible)
        Button11.Visible = True
        TextBox1.Visible = False
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        ' make Text visible (and Image invisible)
        Button11.Visible = False
        TextBox1.Visible = True
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        ' make Image visible (and text invisible)
        Button12.Visible = True
        TextBox2.Visible = False
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        ' make Text visible (and Image invisible)
        Button12.Visible = False
        TextBox2.Visible = True
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        ' make Image visible (and text invisible)
        Button13.Visible = True
        TextBox3.Visible = False
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        ' make Text visible (and Image invisible)
        Button13.Visible = False
        TextBox3.Visible = True
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        ' make Image visible (and text invisible)
        Button14.Visible = True
        TextBox4.Visible = False
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        ' make Text visible (and Image invisible)
        Button14.Visible = False
        TextBox4.Visible = True
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        ' make Image visible (and text invisible)
        Button15.Visible = True
        TextBox5.Visible = False
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        ' make Text visible (and Image invisible)
        Button15.Visible = False
        TextBox5.Visible = True
    End Sub

    Private Sub HelpToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem1.Click

        Dim HlpMsg As String

        HlpMsg = "  "
        HlpMsg = HlpMsg & " Multi-Clip is a very simple tool with 5 boxes into which text and " & vbCrLf
        HlpMsg = HlpMsg & "image data from the system clipboard can be pasted, and from which " & vbCrLf
        HlpMsg = HlpMsg & "data can be copied back to the system clipboard." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  Eack box shows its contents in a small window - the actual " & vbCrLf
        HlpMsg = HlpMsg & "contents may be much larger so the small windows are just to " & vbCrLf
        HlpMsg = HlpMsg & "help identify the box contents. " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  Paste :" & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  The paste action duplicates the data from the clipboard into the box." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  If the clipboard contains complex data such as formatted text from " & vbCrLf
        HlpMsg = HlpMsg & "MS Word only the text within it is duplcated to the box - the " & vbCrLf
        HlpMsg = HlpMsg & "fomatting is ignored. " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  The clipboard contents are unchanged - so still include any formatting." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  If the clipboard contains image data it is duplicated to the active " & vbCrLf
        HlpMsg = HlpMsg & "image box and a thumbnail is displayed -  again the clipboard remains " & vbCrLf
        HlpMsg = HlpMsg & "unchanged. " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  When an image is first pasted into each box the image mode becomes " & vbCrLf
        HlpMsg = HlpMsg & "active and two small buttons are displayed which will switch the " & vbCrLf
        HlpMsg = HlpMsg & "active mode of the box to either text or image mode." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  If the clipboard contents are too complex and neither plain text " & vbCrLf
        HlpMsg = HlpMsg & "nor an image can be derived from it then an error message is displayed " & vbCrLf
        HlpMsg = HlpMsg & "- and again the clipboard's contents are left unchanged." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  Copy :" & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  The copy action copies the active box contents into the clipboard. " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  Each Multi-Clip box can hold both text and an image - the starting " & vbCrLf
        HlpMsg = HlpMsg & "active mode is text.  " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  Pasting text into a box does not affect any image already there " & vbCrLf
        HlpMsg = HlpMsg & "and pasting an image does not affect any text already there." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  The active mode, is which ever is visible, text or image. " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  The small text windows which show the box's text contents do allow " & vbCrLf
        HlpMsg = HlpMsg & "simple editing - but this is not their primary purpose - so there are " & vbCrLf
        HlpMsg = HlpMsg & "no additional text editing controls. " & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "  The image data is held unchanged (the thumbnail is just an identifier). " & vbCrLf
        HlpMsg = HlpMsg & "There are no image editing or conversion functions built into Multi-Clip." & vbCrLf
        HlpMsg = HlpMsg & " " & vbCrLf
        ' HlpMsg = HlpMsg & "  Planned development is to add support for some more common formatted " & vbCrLf
        ' HlpMsg = HlpMsg & "data so Multi-Clip will preserve formatting for common data sources such " & vbCrLf
        ' HlpMsg = HlpMsg & "as MS Word, MS Excel and HTML data. " & vbCrLf
        ' HlpMsg = HlpMsg & " " & vbCrLf
        HlpMsg = HlpMsg & "V1.02 - Jan 2016." & vbCrLf

        MsgBox(HlpMsg, 0, "Multi-Clip Help")

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click


        MsgBox("Multi-Clip was writtten in April 2012.", 1, "About Multi-Clip")
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Dim Response As Integer

        Response = MsgBox("Exit Multi-Clip - are you sure ?", 4, "Multi-Clip Exit")
        ' Yes = 6, No = 7 
        'MsgBox(Respnse)

        If Response = 6 Then
            'MsgBox("Exiting")
            End
        End If
        'MsgBox("Not exiting")

    End Sub


    Private Sub Form1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel

    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        ' this button used to check the clipboard contents and set flags for each positive data type.

        ' TextBox6.Text = CountClipboardFormats() - doesn't work in VB ?

        ' look for numeric Clipboard Formats
        TextBox6.Text = "NO"
        FormatCounter = "CF_TEXT"

        

            If My.Computer.Clipboard.ContainsData(FormatCounter) Then
                TextBox6.Text = FormatCounter
            End If






        ' Check for Text
        ' CheckBox1.Visible = False
        CheckBox1.Checked = False
        If My.Computer.Clipboard.ContainsText() Then
            CheckBox1.Visible = True
            CheckBox1.Checked = True
        End If

        ' Check for Image
        ' CheckBox2.Visible = False
        CheckBox2.Checked = False
        If My.Computer.Clipboard.ContainsImage() Then
            CheckBox2.Visible = True
            CheckBox2.Checked = True
        End If

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class