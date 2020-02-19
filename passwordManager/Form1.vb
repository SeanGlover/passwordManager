Option Explicit On
Option Strict On
Imports Controls
Imports Gma.System.MouseKeyHook
Public Class Form1
    Friend WithEvents TypeTabs As New Tabs With {.Dock = DockStyle.Fill,
        .MouseOverSelection = True,
        .SelectedTabColor = Color.Purple,
        .UserCanAdd = True,
        .UserCanReorder = True,
        .Alignment = TabAlignment.Top,
        .Multiline = True}
    Private ReadOnly TLP_AddTab As New TableLayoutPanel With {.ColumnCount = 2,
        .RowCount = 1,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .BorderStyle = BorderStyle.Fixed3D,
        .Size = New Size(236, 36)}
    Private WithEvents CategoryPicture As New PictureBox With {.Dock = DockStyle.Fill,
        .BackgroundImage = My.Resources.NoImage,
        .BackgroundImageLayout = ImageLayout.Center,
        .Margin = New Padding(0)}
    Private WithEvents CategoryCombo As New ImageCombo With {.HintText = "Add new category",
        .Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .Font = New Font("Century Gothic", 14, FontStyle.Bold)}
    Private WithEvents PanelCategories As New TableLayoutPanel With {.RowCount = 10,
            .ColumnCount = 10,
            .BorderStyle = BorderStyle.Fixed3D,
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
            .Dock = DockStyle.Fill}
    Friend Shared ReadOnly InvisibleForm As New Form With {
        .FormBorderStyle = FormBorderStyle.None,
        .Size = New Size(150, 150),
        .BackgroundImageLayout = ImageLayout.Center,
        .ShowInTaskbar = False
        }
    Private ReadOnly ControlsImages As New Dictionary(Of String, Image)(MyImages(Nothing))
    Friend ReadOnly NoImageString As String = ImageToBase64(My.Resources.NoImage)
    Private Categories As CategoryCollection

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        WindowState = FormWindowState.Maximized
        Icon = My.Resources.Lock

        With TLP_AddTab
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 36})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 200})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
            With .Controls
                .Add(CategoryPicture, 0, 0)
                .Add(CategoryCombo, 1, 0)
            End With
        End With
        TLP.SetSize(TLP_AddTab)

        Controls.Add(TypeTabs)

        Categories = New CategoryCollection
        For Each category In Categories
            With TypeTabs.TabPages
                Dim categoryTab As Tab = .Add(category.Name)
                categoryTab.Controls.Add(category.Panel)
                categoryTab.Image = category.TabImage
            End With
        Next

        With InvisibleForm
            .Show(Me)
            .Visible = False
        End With

    End Sub

    Private Sub TabControl_AddTabClicked(sender As Object, e As TabsEventArgs) Handles TypeTabs.TabClicked

        If e.InZone = Tabs.Zone.Add Then
            'Add new Category
            InvisibleForm.Controls.Clear()
            InvisibleForm.Controls.Add(TLP_AddTab)
            TLP_AddTab.Location = New Point(32, 32)
            Dim questionsRectangle As New Rectangle(New Point(0, 0), TLP_AddTab.Size)
            questionsRectangle.Inflate(32, 32)
            With InvisibleForm
                .Size = questionsRectangle.Size
                .BackgroundImage = FormBackImage(questionsRectangle.Size)
                .Visible = True
            End With

        ElseIf e.InZone = Tabs.Zone.Close Then
            'Drop existing Category
            Using message As New Prompt
                If message.Show("Are you sure?", "Removing this category will remove all entries", Prompt.IconOption.YesNo) = DialogResult.Yes Then
                    Dim oldCategory As EntryCollection = Categories.Item(Trim(e.InTab.Text))
                    If oldCategory IsNot Nothing Then Categories.Remove(oldCategory)
                    TypeTabs.TabPages.Remove(e.InTab)
                    Categories.Save()
                End If
            End Using

        ElseIf e.InZone = Tabs.Zone.Image Or e.InZone = Tabs.Zone.Text Then
            If Clipboard.ContainsImage Then
                Dim clipImage As Image = Clipboard.GetImage()
                Dim clickedCategory As EntryCollection = Categories.Item(Trim(e.InTab.Text))
                clickedCategory.TabImageString = ImageToBase64(clipImage)
                e.InTab.Image = clickedCategory.TabImage
                Categories.Save()

            Else
                e.InTab.Image = Nothing
                Dim clickedCategory As EntryCollection = Categories.Item(Trim(e.InTab.Text))
                clickedCategory.TabImageString = Nothing
                Categories.Save()

            End If

        End If

    End Sub
    Private Sub NewTabImage_Click() Handles CategoryPicture.Click

        If Clipboard.ContainsImage Then
            Dim clipImage As Image = Clipboard.GetImage()
            CategoryPicture.InitialImage = clipImage
            CategoryPicture.BackgroundImage = ResizeImage(clipImage, CategoryPicture.Size)

        Else
            CategoryPicture.BackgroundImage = My.Resources.NoImage
        End If

    End Sub
    Private Sub NewTabCombo_ValueSubmitted(sender As Object, e As ImageComboEventArgs) Handles CategoryCombo.ValueSubmitted

        If CategoryCombo.Text.Any Then
            Dim newTab As Tab = TypeTabs.TabPages.Add(CategoryCombo.Text)
            newTab.Image = ResizeImage(CategoryPicture.InitialImage, 20, 20)
            Dim newCategory = Categories.Add(New EntryCollection(CategoryCombo.Text))
            newCategory.TabImageString = ImageToBase64(CategoryPicture.InitialImage)
            newTab.Controls.Add(newCategory.Panel)
        End If
        Categories.Save()

    End Sub

    Friend Shared Function FormBackImage(formSize As Size) As Bitmap

        'OuterSquare- pen
        'InnerRectangle - solid brush

        Dim bmp As New Bitmap(formSize.Width, formSize.Height)
        Dim defaultColor As Color = Color.Purple

        Using graphics As Graphics = Graphics.FromImage(bmp)
            With graphics
                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                Dim rect As New Rectangle(New Point(0, 0), formSize)
                .FillRectangle(Brushes.Maroon, rect)
                .Clear(Color.Transparent)
                Using borderPen As New Pen(Brushes.Purple, 12)
                    .DrawRectangle(borderPen, rect)
                    rect.Inflate(-16, -16)
                    Using fillBrush As New SolidBrush(Color.Gainsboro)
                        .FillRectangle(Brushes.Gainsboro, rect)
                    End Using
                End Using
            End With
        End Using
        bmp.MakeTransparent(Color.Maroon)
        Return bmp

    End Function

End Class

Public Class CategoryCollection
    Inherits List(Of EntryCollection)
    Public Sub New()

        Dim pmString As String
        With My.Settings
            '.encryptedValues = "Banking╠BMO0§§bmo.com§abc§def►What is my favorite color0§Blue■What is my favorite color1§Blue■What is my favorite color2§Blue■What is my favorite color3§Blue■What is my favorite color4§Blue■What is my favorite color5§Blue●BMO1§§bmo.com§abc§def►What is my favorite color0§Blue■What is my favorite color1§Blue■What is my favorite color2§Blue■What is my favorite color3§Blue■What is my favorite color4§Blue■What is my favorite color5§Blue●BMO2§§bmo.com§abc§def►What is my favorite color0§Blue■What is my favorite color1§Blue■What is my favorite color2§Blue■What is my favorite color3§Blue■What is my favorite color4§Blue■What is my favorite color5§Blue●BMO3§§bmo.com§abc§def►What is my favorite color0§Blue■What is my favorite color1§Blue■What is my favorite color2§Blue■What is my favorite color3§Blue■What is my favorite color4§Blue■What is my favorite color5§Blue●BMO4§§bmo.com§abc§def►What is my favorite color0§Blue■What is my favorite color1§Blue■What is my favorite color2§Blue■What is my favorite color3§Blue■What is my favorite color4§Blue■What is my favorite color5§Blue●BMO5§§bmo.com§abc§def►What is my favorite color0§Blue■What is my favorite color1§Blue■What is my favorite color2§Blue■What is my favorite color3§Blue■What is my favorite color4§Blue■What is my favorite color5§Blue╬Work╠CCS0§§bmo.com§abc§def►What is my name0§Sean■What is my name1§Sean■What is my name2§Sean■What is my name3§Sean■What is my name4§Sean■What is my name5§Sean●CCS1§§bmo.com§abc§def►What is my name0§Sean■What is my name1§Sean■What is my name2§Sean■What is my name3§Sean■What is my name4§Sean■What is my name5§Sean●CCS2§§bmo.com§abc§def►What is my name0§Sean■What is my name1§Sean■What is my name2§Sean■What is my name3§Sean■What is my name4§Sean■What is my name5§Sean●CCS3§§bmo.com§abc§def►What is my name0§Sean■What is my name1§Sean■What is my name2§Sean■What is my name3§Sean■What is my name4§Sean■What is my name5§Sean●CCS4§§bmo.com§abc§def►What is my name0§Sean■What is my name1§Sean■What is my name2§Sean■What is my name3§Sean■What is my name4§Sean■What is my name5§Sean●CCS5§§bmo.com§abc§def►What is my name0§Sean■What is my name1§Sean■What is my name2§Sean■What is my name3§Sean■What is my name4§Sean■What is my name5§Sean"
            '.Save()
            pmString = .encryptedValues
        End With
        Dim collectionTypes As New List(Of String)(Split(pmString, "╬"))
        collectionTypes.Sort()

        For Each collectionType In collectionTypes
            Dim typeElements As New List(Of String)(Split(collectionType, "╠"))
            Dim typeNameImage As String = typeElements.First
            Dim elementsNameImage As String() = Split(typeNameImage, "₪")
            Dim typeName As String = elementsNameImage.First
            Dim typeImage As String = If(elementsNameImage.Length = 2, elementsNameImage.Last, String.Empty)
            Dim typeString As String = typeElements.Last
            Dim newCollectionType As New EntryCollection(typeName) With {.TabImageString = typeImage}
            Dim entryList As New List(Of String)(Split(typeString, "●"))
            entryList.Sort(Function(f1, f2)
                               Dim Level1 = String.Compare(Split(f1, Delimiter)(0).ToUpperInvariant, Split(f2, Delimiter)(0).ToUpperInvariant, StringComparison.Ordinal)
                               If Level1 <> 0 Then
                                   Return Level1
                               Else
                                   Dim Level2 = String.Compare(Split(f1, Delimiter)(2).ToUpperInvariant, Split(f2, Delimiter)(2).ToUpperInvariant, StringComparison.Ordinal)
                                   Return Level2
                               End If
                           End Function)

            For Each entryItem In entryList
                If entryItem.Any Then
                    Dim entryQuestions As New List(Of String)(Split(entryItem, "►"))
                    Dim entryElements As New List(Of String)(Split(entryQuestions.First, Delimiter))
                    Dim entryName As String = entryElements.First
                    Dim entryImageString As String = entryElements(1)
                    Dim logo As Image
                    Try
                        logo = Base64ToImage(entryImageString)
                    Catch ex As Exception
                        logo = My.Resources.NoImage
                        entryImageString = ImageToBase64(logo)
                    End Try
                    Dim entryNickname As String = entryElements(2)
                    Dim entryURL As String = entryElements(3)
                    Dim entryUID As String = entryElements(4)
                    Dim entryPWD As String = entryElements.Last
                    Dim newEntry As Entry = newCollectionType.Add(entryName,
                                                                logo,
                                                                entryImageString,
                                                                entryNickname,
                                                                entryURL,
                                                                entryUID,
                                                                entryPWD)
                    Dim securityQuestions As New List(Of String)(Split(entryQuestions.Last, BlackOut))
                    securityQuestions.Sort()

                    For Each securityQuestion In securityQuestions
                        Dim qaElements As New List(Of String)(Split(securityQuestion, Delimiter))
                        newEntry.SecurityQuestions.Add(qaElements.First, qaElements.Last)
                    Next
                End If
            Next
            Add(newCollectionType)
        Next

    End Sub

    Public Shadows Function Item(categoryName As String) As EntryCollection

        Dim matchingCollection As EntryCollection = Nothing
        For Each collection In Me
            If collection.Name.ToUpperInvariant = categoryName.ToUpperInvariant Then
                matchingCollection = collection
                Exit For
            End If
        Next
        Return matchingCollection

    End Function
    Public Shadows Function Add(collectionType As EntryCollection) As EntryCollection

        If collectionType IsNot Nothing Then
            collectionType.Parent_ = Me
            MyBase.Add(collectionType)
        End If
        Return collectionType

    End Function
    Friend Sub Save()

        With My.Settings
            .encryptedValues = ToString()
            .Save()
            Using sw As New IO.StreamWriter(Desktop & "\pm.txt")
                sw.Write(ToString)
            End Using
        End With

    End Sub
    Public Overrides Function ToString() As String
        Return Microsoft.VisualBasic.Join((From ec As EntryCollection In Me Select ec.ToString & String.Empty).ToArray, "╬")
    End Function
End Class
Public Class EntryCollection
    Inherits List(Of Entry)
    Private WithEvents BaseForm As Form = Form1
    Private WithEvents TabControl As Tabs = Form1.TypeTabs
    Friend ReadOnly MH As New SurroundingClass

    Public Enum Type
        Other
        Banking
        Work
        Personal
        Payments
    End Enum
    Friend Parent_ As CategoryCollection
    Public ReadOnly Property Parent As CategoryCollection
        Get
            Return Parent_
        End Get
    End Property
    Friend ReadOnly Panel As New TableLayoutPanel With {.RowCount = 1,
    .ColumnCount = 1,
    .BorderStyle = BorderStyle.FixedSingle,
    .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
    .Margin = New Padding(0),
    .BackColor = Color.Blue}
    Friend WithEvents AddButton As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .BackColor = Color.Gainsboro,
        .BackgroundImage = My.Resources.Button_Light,
        .Text = "A D D",
        .BackgroundImageLayout = ImageLayout.Stretch}
    Friend WithEvents SubmitButton As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .BackColor = Color.Gainsboro,
        .BackgroundImage = My.Resources.Button_Light,
        .Text = "S U B M I T   C H A N G E S",
        .BackgroundImageLayout = ImageLayout.Stretch}
    Private ReadOnly Property Combonames As String
        Get
            Dim Entries As New List(Of String)
            For Each entry In Me
                Entries.Add(Microsoft.VisualBasic.Join({entry.ComboName.Name.ToUpperInvariant,
                                                     entry.ButtonImage.Name,
                                                     entry.ComboNickName.Name.ToUpperInvariant,
                                                     entry.ComboURL.Name.ToUpperInvariant,
                                                     entry.ComboUID.Name.ToUpperInvariant,
                                                     entry.ComboPWD.Name}, Delimiter))
            Next
            Return Microsoft.VisualBasic.Join(Entries.ToArray, BlackOut)
        End Get
    End Property
    Private ReadOnly Property Combotexts As String
        Get
            Dim Entries As New List(Of String)
            For Each entry In Me
                Entries.Add(Microsoft.VisualBasic.Join({entry.ComboName.Text.ToUpperInvariant,
                                                     entry.ButtonImage.Text,
                                                     entry.ComboNickName.Text.ToUpperInvariant,
                                                     entry.ComboURL.Text.ToUpperInvariant,
                                                     entry.ComboUID.Text.ToUpperInvariant,
                                                     entry.ComboPWD.Text}, Delimiter))
            Next
            Return Microsoft.VisualBasic.Join(Entries.ToArray, BlackOut)
        End Get
    End Property
    Friend ReadOnly Property CollectionType As Type
        Get
            Return ParseEnum(Of Type)(Name)
        End Get
    End Property
    Friend Property Name As String
    Friend Property TabImageString As String
    Friend ReadOnly Property TabImage As Image
        Get
            Return If(TabImageString.Any, ResizeImage(Base64ToImage(TabImageString), 20, 20), Nothing)
        End Get
    End Property

    Public Sub New(collectionName As String)

        Name = collectionName
        AddButton.Text = Microsoft.VisualBasic.Join({"ADD", Name, "item"})
        Dim addWidth As Integer
        Using buttonFont As New Font("Century Gothic", 14, FontStyle.Bold)
            AddButton.Font = buttonFont
            SubmitButton.Font = buttonFont
            addWidth = 16 + TextRenderer.MeasureText(AddButton.Text, buttonFont).Width
        End Using
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = WorkingArea.Width - 70})
        Panel.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
        Dim tlpAddSubmit As New TableLayoutPanel With {.RowCount = 1, .ColumnCount = 2, .Dock = DockStyle.Fill}
        With tlpAddSubmit
            .Margin = New Padding(0)
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            .BorderStyle = BorderStyle.None
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = addWidth})
            .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = Panel.ColumnStyles(0).Width - addWidth})
            .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
            .Controls.Add(AddButton, 0, 0)
            .Controls.Add(SubmitButton, 1, 0)
        End With
        TLP.SetSize(tlpAddSubmit)
        Panel.Controls.Add(tlpAddSubmit, 0, 0)
        TLP.SetSize(Panel)
        AddHandler TabControl.SelectedIndexChanged, AddressOf TabControl_IndexChanged

    End Sub

    Public Shadows Function Contains(name As String, userid As String) As Boolean

        Dim matches As New List(Of Entry)
        For Each entryItem As Entry In Me
            If entryItem.ComboName.Text.ToUpperInvariant = name.ToUpperInvariant And entryItem.ComboUID.Text.ToUpperInvariant = userid.ToUpperInvariant Then matches.Add(entryItem)
        Next
        Return matches.Any

    End Function
    Public Shadows Function Add(Name As String, backImage As Image, backImageString As String, Nickname As String, Address As String, UserID As String, UserPWD As String) As Entry

        If Contains(Name, UserID) Then
            Return Nothing
        Else
            Panel.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 48})
            Dim newEntry As New Entry()
            With newEntry
                .ComboName.Text = Name
                .ButtonImage.BackgroundImage = backImage
                .ComboNickName.Text = Nickname
                .ComboURL.Text = Address
                .ComboURL.Image = If(backImageString = Form1.NoImageString, Nothing, ResizeImage(backImage, 24, 24))
                .ComboUID.Text = UserID
                .ComboPWD.Text = UserPWD
                .ComboName.Name = Name
                .ButtonImage.Name = backImageString
                .ComboNickName.Name = Nickname
                .ComboURL.Name = Address
                .ComboUID.Name = UserID
                .ComboPWD.Name = UserPWD
                .Parent_ = Me
                .ChangeName(Name)
                AddHandler .ButtonRemove.Click, AddressOf Entry_Remove
                AddHandler .Panel.Controls.OfType(Of PictureBox).ToList.First.Click, AddressOf Element_Changed
                AddHandler .ComboURL.ImageClicked, AddressOf URL_ImageClicked
                .Panel.Controls.OfType(Of ImageCombo).ToList.ForEach(Function(Item) As ImageCombo
                                                                         AddHandler Item.TextChanged, AddressOf Element_Changed
                                                                         Return Nothing
                                                                     End Function)
                Panel.Controls.Add(.Panel, 0, Panel.Controls.Count)
            End With
            TLP.SetSize(Panel)

            MyBase.Add(newEntry)
            Return newEntry
        End If

    End Function
    Public Shadows Function Remove(oldEntry As Entry) As Entry

        If oldEntry IsNot Nothing Then
            Dim rowcolumnIndex As TableLayoutPanelCellPosition = Panel.GetCellPosition(oldEntry.Panel)
            Panel.Controls.Remove(oldEntry.Panel)
            Panel.RowStyles(rowcolumnIndex.Row).Height = 0
            TLP.SetSize(Panel)
            oldEntry.Panel.Controls.OfType(Of ImageCombo)().ToList().ForEach(Function(Item) As ImageCombo
                                                                                 RemoveHandler Item.TextChanged, AddressOf Element_Changed
                                                                                 Return Nothing
                                                                             End Function)
            RemoveHandler oldEntry.ButtonRemove.Click, AddressOf Entry_Remove
            RemoveHandler oldEntry.Panel.Controls.OfType(Of PictureBox).ToList.First.Click, AddressOf Element_Changed
            MyBase.Remove(oldEntry)
        End If
        Return oldEntry

    End Function
    Private Sub Element_Changed(sender As Object, e As EventArgs)

        SubmitButton.BackgroundImage = If(Combonames = Combotexts, My.Resources.Button_Light, My.Resources.Button_Bright)

    End Sub
    Private Sub Changes_Submitted(sender As Object, e As EventArgs) Handles SubmitButton.Click

        For Each entryItem As Entry In Me
            Dim imageCombos = entryItem.Panel.Controls.OfType(Of ImageCombo)().ToList()
            imageCombos.ForEach(Function(Item) As ImageCombo
                                    Item.Name = Item.Text
                                    Return Nothing
                                End Function)
            Dim itemName As String = imageCombos.First.Name
            entryItem.ChangeName(itemName)
            Dim buttonImage As PictureBox = entryItem.Panel.Controls.OfType(Of PictureBox).ToList.First
            buttonImage.Text = buttonImage.Name
        Next
        Parent.Save()
        SubmitButton.BackgroundImage = My.Resources.Button_Light

    End Sub
    Private Sub Entry_Remove(sender As Object, e As EventArgs)

        With DirectCast(sender, Button)
            Dim buttonPanel As TableLayoutPanel = DirectCast(.Parent, TableLayoutPanel)
            Dim whichEntry As Entry = (From p In Me Where p.Panel Is buttonPanel).First
            Remove(whichEntry)
        End With
        SubmitButton.BackgroundImage = My.Resources.Button_Bright

    End Sub
    Private Sub AddButton_Clicked() Handles AddButton.Click
        Add(String.Empty, My.Resources.NoImage, ImageToBase64(My.Resources.NoImage), String.Empty, String.Empty, String.Empty, String.Empty)
    End Sub
    Private Sub URL_ImageClicked(sender As Object, e As ImageComboEventArgs)
        With DirectCast(sender, ImageCombo)
            Process.Start("chrome.exe", .Text)
        End With
    End Sub

    Friend Sub QuestionButton_Enter(mouseEntry As Entry)

        If Not MH.Hooked Then
            With Form1.InvisibleForm
                Dim questionsRectangle As Rectangle = mouseEntry.SecurityQuestions.Panel.ClientRectangle
                questionsRectangle.Inflate(32, 32)
                .BackgroundImage = Form1.FormBackImage(questionsRectangle.Size)
                .Size = questionsRectangle.Size
                .Controls.Add(mouseEntry.SecurityQuestions.Panel)
                mouseEntry.SecurityQuestions.Panel.Location = New Point(32, 32)
                .Location = New Point(20, mouseEntry.ButtonQuestions.PointToScreen(New Point(0, 0)).Y)
                .Visible = True
                mouseEntry.SecurityQuestions.AddButton.Image = ResizeImage(mouseEntry.ButtonImage.BackgroundImage, 30, 30)
                AddHandler MH.Moused, AddressOf Hook_MouseDown
            End With
        End If

    End Sub
    Friend Sub QuestionButton_Clicked(mouseEntry As Entry)
        MH.Subscribe()
    End Sub
    Friend Sub QuestionButton_Leave()

        If Not MH.Hooked Then
            With Form1.InvisibleForm
                .Controls.Clear()
                .Visible = False
            End With
        End If

    End Sub
    Private Sub Hook_MouseDown(sender As Object, e As MouseEventExtArgs)

        If e.Button = MouseButtons.Left Then
            If Not CursorOverControl(Form1.InvisibleForm) Then
                QuestionButton_Leave()
                MH.Unsubscribe()
            End If
        End If

    End Sub
    Private Sub ParentForm_Closing() Handles BaseForm.Closing
        MH.Unsubscribe()
    End Sub
    Private Sub TabControl_IndexChanged(sender As Object, e As EventArgs)

        Form1.InvisibleForm.Visible = False
        MH.Unsubscribe()

    End Sub
    Public Overrides Function ToString() As String
        Return Microsoft.VisualBasic.Join({Microsoft.VisualBasic.Join({Name, TabImageString}, "₪"), Microsoft.VisualBasic.Join((From entry In Me Select entry.ToString & String.Empty).ToArray, "●")}, "╠")
    End Function
End Class

Public Class Entry
    'Entry {[X], Name, Nickname, URL, UID, PWD, QuestionCollection}
    Friend Parent_ As EntryCollection
    Public ReadOnly Property Parent As EntryCollection
        Get
            Return Parent_
        End Get
    End Property
    Public ReadOnly Property SecurityQuestions As QuestionCollection
    Friend ReadOnly Panel As New TableLayoutPanel With {.RowCount = 1,
        .ColumnCount = 8,
        .Size = New Size(WorkingArea.Width, 32),
        .BorderStyle = BorderStyle.Fixed3D,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .Margin = New Padding(0)}
    Friend WithEvents ButtonRemove As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .FlatStyle = FlatStyle.Flat,
        .BackgroundImage = My.Resources.RemoveGrey.ToBitmap,
        .BackgroundImageLayout = ImageLayout.Center,
        .BackColor = Color.GhostWhite}
    Friend WithEvents ButtonImage As New PictureBox With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .BackgroundImage = My.Resources.NoImage,
        .BackgroundImageLayout = ImageLayout.Stretch,
        .BackColor = Color.GhostWhite,
        .BorderStyle = BorderStyle.Fixed3D}
    Friend WithEvents ButtonQuestions As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .FlatStyle = FlatStyle.Flat,
        .BackgroundImage = My.Resources.Question.ToBitmap,
        .BackgroundImageLayout = ImageLayout.Center}
    Friend ReadOnly ComboName As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Name"}
    Friend ReadOnly ComboNickName As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Nickname"}
    Friend ReadOnly ComboURL As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "URL"}
    Friend ReadOnly ComboUID As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "User ID"}
    Friend ReadOnly ComboPWD As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Password"}
    Public Sub New()

        SecurityQuestions = New QuestionCollection(Me)
        Using comboFont As New Font("Century Gothic", 12, FontStyle.Bold)
            ComboName.Font = comboFont
            ComboNickName.Font = comboFont
            'ComboURL.Font = comboFont
            ComboUID.Font = comboFont
            ComboPWD.Font = comboFont
        End Using
        Dim xButtonWidth As Integer = 40
        Dim remainingWidth As Integer = Panel.Width - xButtonWidth
        Dim namesWidth As Integer = CInt(remainingWidth * 0.12)
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = xButtonWidth})
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = xButtonWidth})
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = namesWidth}) '15
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = namesWidth}) '15=30
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = namesWidth * 2}) '30=60
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = namesWidth}) '15=75
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = namesWidth}) '15=90
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = 32})
        Panel.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute,
                               .Height = Panel.Height})
        Panel.Controls.Add(ButtonRemove, 0, 0)
        Panel.Controls.Add(ButtonImage, 1, 0)
        Panel.Controls.Add(ComboName, 2, 0)
        Panel.Controls.Add(ComboNickName, 3, 0)
        Panel.Controls.Add(ComboURL, 4, 0)
        Panel.Controls.Add(ComboUID, 5, 0)
        Panel.Controls.Add(ComboPWD, 6, 0)
        Panel.Controls.Add(ButtonQuestions, 7, 0)
        TLP.SetSize(Panel)

    End Sub

    Private Sub ButtonImage_Click() Handles ButtonImage.Click

        If Clipboard.ContainsImage Then
            Dim clipImage As Image = Clipboard.GetImage()
            ButtonImage.BackgroundImage = clipImage
            ButtonImage.Name = ImageToBase64(clipImage)
            ComboURL.Image = ResizeImage(clipImage, 24, 24)
        Else
            ButtonImage.BackgroundImage = My.Resources.NoImage
            ButtonImage.Name = String.Empty
            ComboURL.Image = Nothing
        End If

    End Sub
    Private Sub ButtonRemove_Enter() Handles ButtonRemove.MouseEnter
        ButtonRemove.BackgroundImage = My.Resources.Remove.ToBitmap
    End Sub
    Private Sub ButtonRemove_Leave() Handles ButtonRemove.MouseLeave
        ButtonRemove.BackgroundImage = My.Resources.RemoveGrey.ToBitmap
    End Sub
    Private Sub ButtonQuestions_Enter() Handles ButtonQuestions.MouseEnter
        Parent.QuestionButton_Enter(Me)
    End Sub
    Private Sub Questions_Clicked(sender As Object, e As EventArgs) Handles ButtonQuestions.Click
        Parent.QuestionButton_Clicked(Me)
    End Sub
    Private Sub ButtonQuestions_Leave() Handles ButtonQuestions.MouseLeave
        Parent.QuestionButton_Leave()
    End Sub
    Friend Sub ChangeName(newName As String)
        SecurityQuestions.ChangeName(newName)
    End Sub
    Public Overrides Function ToString() As String
        Return Join({Join({ComboName.Name, ButtonImage.Name, ComboNickName.Name, ComboURL.Name, ComboUID.Name, ComboPWD.Name}, Delimiter), Join((From qa In SecurityQuestions Select qa.ToString & String.Empty).ToArray, BlackOut)}, "►")
    End Function
End Class

Public Class QuestionCollection
    Inherits List(Of QA)
    Private Const RowHeight As Integer = 40
    Friend ReadOnly Property Parent As Entry
    Friend ReadOnly Panel As New TableLayoutPanel With {.RowCount = 1,
        .ColumnCount = 1,
        .BorderStyle = BorderStyle.FixedSingle,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.None,
        .Margin = New Padding(0),
        .BackColor = Color.Red}
    Friend WithEvents AddButton As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .BackColor = Color.Gainsboro,
        .BackgroundImage = My.Resources.Button_Light,
        .Text = "A D D",
        .BackgroundImageLayout = ImageLayout.Stretch,
        .Image = My.Resources.Question.ToBitmap,
        .TextImageRelation = TextImageRelation.TextBeforeImage}
    Friend WithEvents SubmitButton As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .BackColor = Color.Gainsboro,
        .BackgroundImage = My.Resources.Button_Light,
        .Text = "S U B M I T   C H A N G E S",
        .BackgroundImageLayout = ImageLayout.Stretch}
    Private ReadOnly Property QAnames As String
        Get
            Dim QAs As New List(Of String)
            For Each qa In Me
                QAs.Add(Microsoft.VisualBasic.Join({qa.ComboQuestion.Name.ToUpperInvariant, qa.ComboAnswer.Name.ToUpperInvariant}, Delimiter))
            Next
            Return Microsoft.VisualBasic.Join(QAs.ToArray, BlackOut)
        End Get
    End Property
    Private ReadOnly Property QAtexts As String
        Get
            Dim QAs As New List(Of String)
            For Each qa In Me
                QAs.Add(Microsoft.VisualBasic.Join({qa.ComboQuestion.Text.ToUpperInvariant, qa.ComboAnswer.Text.ToUpperInvariant}, Delimiter))
            Next
            Return Microsoft.VisualBasic.Join(QAs.ToArray, BlackOut)
        End Get
    End Property

    Public Sub New(parentEntry As Entry)

        Parent = parentEntry
        If parentEntry IsNot Nothing Then
            AddButton.Text = Microsoft.VisualBasic.Join({"ADD", parentEntry.ComboName.Name, "item"})
            AddButton.Image = If(SameImage(Parent.ButtonImage.BackgroundImage, My.Resources.NoImage), My.Resources.Question.ToBitmap, Parent.ButtonImage.BackgroundImage)
            Dim addWidth As Integer
            Using buttonFont As New Font("Century Gothic", 14, FontStyle.Bold)
                AddButton.Font = buttonFont
                SubmitButton.Font = buttonFont
                addWidth = AddButton.Image.Width + 16 + TextRenderer.MeasureText(AddButton.Text, buttonFont).Width
            End Using
            Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = 800})
            Panel.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = RowHeight})
            Dim tlpAddSubmit As New TableLayoutPanel With {.RowCount = 1, .ColumnCount = 2, .Dock = DockStyle.Fill}
            With tlpAddSubmit
                .Margin = New Padding(0)
                .CellBorderStyle = TableLayoutPanelCellBorderStyle.None
                .BackColor = Color.GhostWhite
                .BorderStyle = BorderStyle.None
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = addWidth})
                .ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute, .Width = Panel.ColumnStyles(0).Width - addWidth})
                .RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = 36})
                .Controls.Add(AddButton, 0, 0)
                .Controls.Add(SubmitButton, 1, 0)
            End With
            Panel.Controls.Add(tlpAddSubmit, 0, 0)
        End If

    End Sub

    Friend Sub ChangeName(newName As String)

        AddButton.Text = Microsoft.VisualBasic.Join({"ADD", newName, "question"})
        Dim addWidth As Integer
        Using buttonFont As New Font("Century Gothic", 14, FontStyle.Bold)
            AddButton.Font = buttonFont
            SubmitButton.Font = buttonFont
            addWidth = SystemIcons.Question.Width + 16 + TextRenderer.MeasureText(AddButton.Text, buttonFont).Width
        End Using
        Dim tlpAddSubmit As TableLayoutPanel = DirectCast(Panel.Controls(0), TableLayoutPanel)
        With tlpAddSubmit.ColumnStyles
            .Item(0).Width = addWidth
            .Item(1).Width = Panel.ColumnStyles(0).Width - addWidth
        End With

    End Sub
    Public Shadows Function Contains(question As String) As Boolean

        Dim matches As New List(Of QA)
        For Each questionAnswer As QA In Me
            If questionAnswer.ComboQuestion.Name.ToUpperInvariant = question.ToUpperInvariant Then matches.Add(questionAnswer)
        Next
        Return matches.Any

    End Function
    Public Shadows Function Add(question As String, answer As String) As QA

        If Contains(question) Then
            Return Nothing
        Else
            Dim newQuestionAnswer As New QA
            With newQuestionAnswer
                .ComboQuestion.Text = question
                .ComboQuestion.Name = question
                .ComboAnswer.Text = answer
                .ComboAnswer.Name = answer
                .Parent_ = Me
                Panel.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute, .Height = RowHeight})
                Panel.Controls.Add(.Panel, 0, Panel.Controls.Count)

                ResizePanel()

                AddHandler .ComboQuestion.TextChanged, AddressOf QuestionOrAnswer_TextChanged
                AddHandler .ComboAnswer.TextChanged, AddressOf QuestionOrAnswer_TextChanged
                AddHandler .ComboQuestion.ValueSubmitted, AddressOf Combo_ValueSubmitted
                AddHandler .ComboAnswer.ValueSubmitted, AddressOf Combo_ValueSubmitted
                AddHandler .ButtonRemove.Click, AddressOf Question_Remove
                MyBase.Add(newQuestionAnswer)
            End With
            Return newQuestionAnswer
        End If

    End Function
    Private Sub ResizePanel()

        Dim rowStyles As New List(Of RowStyle)(From rs In Panel.RowStyles Select DirectCast(rs, RowStyle))
        Dim hiddenRowstyles As New List(Of RowStyle)(From rs In rowStyles Where rs.Height = 0)
        For Each rs As RowStyle In hiddenRowstyles
            Panel.RowStyles.Remove(rs)
        Next
        rowStyles = (From rs In Panel.RowStyles Select DirectCast(rs, RowStyle)).ToList
        For Each rs In rowStyles
            rs.Height = RowHeight
        Next
        TLP.SetSize(Panel)
        Panel.Size = New Size(812, Panel.RowStyles.Count * RowHeight)
        Dim questionsRectangle As New Rectangle(New Point(0, 0), Panel.Size)
        questionsRectangle.Inflate(32, 32)
        With Form1.InvisibleForm
            .Size = questionsRectangle.Size
            .BackgroundImage = Form1.FormBackImage(questionsRectangle.Size)
        End With

    End Sub
    Public Shadows Function Remove(oldQuestionAnswer As QA) As QA

        If oldQuestionAnswer IsNot Nothing Then
            With oldQuestionAnswer
                Dim rowcolumnIndex As TableLayoutPanelCellPosition = Panel.GetCellPosition(.Panel)
                Panel.Controls.Remove(.Panel)
                Panel.RowStyles(rowcolumnIndex.Row).Height = 0

                ResizePanel()

                RemoveHandler .ComboQuestion.TextChanged, AddressOf QuestionOrAnswer_TextChanged
                RemoveHandler .ComboAnswer.TextChanged, AddressOf QuestionOrAnswer_TextChanged
                RemoveHandler .ComboQuestion.ValueSubmitted, AddressOf Combo_ValueSubmitted
                RemoveHandler .ComboAnswer.ValueSubmitted, AddressOf Combo_ValueSubmitted
                RemoveHandler .ButtonRemove.Click, AddressOf Question_Remove
            End With
            MyBase.Remove(oldQuestionAnswer)
        End If
        Return oldQuestionAnswer

    End Function
    Private Sub QuestionOrAnswer_TextChanged(sender As Object, e As EventArgs)

        Dim questionCombos = From qa In Me Where qa.ComboQuestion Is sender Or qa.ComboAnswer Is sender
        Dim comboQuestion As ImageCombo = questionCombos.First.ComboQuestion
        Dim comboAnswer As ImageCombo = questionCombos.First.ComboAnswer
        Dim comboText As String = DirectCast(sender, ImageCombo).Text

        Dim matchDelimiter = RegexMatches(comboText, "[=\t\n\r]", System.Text.RegularExpressions.RegexOptions.None)
        If matchDelimiter.Any Then
            RemoveHandler comboQuestion.TextChanged, AddressOf QuestionOrAnswer_TextChanged
            RemoveHandler comboAnswer.TextChanged, AddressOf QuestionOrAnswer_TextChanged

            Dim kvp = System.Text.RegularExpressions.Regex.Split(comboText, "[=\t\n\r]", System.Text.RegularExpressions.RegexOptions.None)
            comboQuestion.Text = kvp.First
            'comboQuestion.Name = kvp.First
            comboAnswer.Text = kvp.Last
            'comboAnswer.Name = kvp.Last

            AddHandler comboQuestion.TextChanged, AddressOf QuestionOrAnswer_TextChanged
            AddHandler comboAnswer.TextChanged, AddressOf QuestionOrAnswer_TextChanged
        End If

        If comboText.Contains("=") Then

        End If
        SubmitButton.BackgroundImage = If(QAnames = QAtexts, My.Resources.Button_Light, My.Resources.Button_Bright)

    End Sub
    Private Sub Combo_ValueSubmitted(sender As Object, e As ImageComboEventArgs)
        Questions_Submitted(Nothing, Nothing)
    End Sub
    Private Sub Questions_Submitted(sender As Object, e As EventArgs) Handles SubmitButton.Click

        For Each qaItem As QA In Me
            Dim imageCombos = qaItem.Panel.Controls.OfType(Of ImageCombo)().ToList()
            imageCombos.ForEach(Function(Item) As ImageCombo
                                    Item.Name = Item.Text
                                    Return Nothing
                                End Function)
        Next
        SubmitButton.BackgroundImage = My.Resources.Button_Light
        Parent.Parent.Parent.Save()

    End Sub
    Private Sub Question_Remove(sender As Object, e As EventArgs)

        With DirectCast(sender, Button)
            Dim buttonPanel As TableLayoutPanel = DirectCast(.Parent, TableLayoutPanel)
            Dim whichQA As QA = (From p In Me Where p.Panel Is buttonPanel).First
            Remove(whichQA)
        End With
        SubmitButton.BackgroundImage = My.Resources.Button_Bright

    End Sub
    Private Sub AddButton_Clicked() Handles AddButton.Click
        Add(String.Empty, String.Empty)
    End Sub
    Public Overrides Function ToString() As String
        Return Microsoft.VisualBasic.Join((From qa In Me Select qa.ToString & String.Empty).ToArray, "║")
    End Function
End Class

Public Class QA
    Friend Parent_ As QuestionCollection
    Public ReadOnly Property Parent As QuestionCollection
        Get
            Return Parent_
        End Get
    End Property
    Friend ReadOnly Panel As New TableLayoutPanel With {.RowCount = 1,
        .ColumnCount = 3,
        .Size = New Size(800, 32),
        .BorderStyle = BorderStyle.Fixed3D,
        .CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset,
        .Margin = New Padding(0)}
    Friend WithEvents ButtonRemove As New Button With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .FlatStyle = FlatStyle.Flat,
        .BackgroundImage = My.Resources.RemoveGrey.ToBitmap,
        .BackgroundImageLayout = ImageLayout.Center,
        .BackColor = Color.GhostWhite}
    Friend WithEvents ComboQuestion As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Question"}
    Friend WithEvents ComboAnswer As New ImageCombo With {.Dock = DockStyle.Fill,
        .Margin = New Padding(0),
        .HintText = "Answer"}
    Public Sub New()

        Using comboFont As New Font("Century Gothic", 12, FontStyle.Bold)
            ComboQuestion.Font = comboFont
            ComboAnswer.Font = comboFont
        End Using
        Dim xButtonWidth As Integer = 32
        Dim icWidth As Integer = CInt((Panel.Width - xButtonWidth) / 2)
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = xButtonWidth})
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = icWidth})
        Panel.ColumnStyles.Add(New ColumnStyle With {.SizeType = SizeType.Absolute,
                               .Width = icWidth})
        Panel.RowStyles.Add(New RowStyle With {.SizeType = SizeType.Absolute,
                               .Height = Panel.Height})
        Panel.Controls.Add(ButtonRemove, 0, 0)
        Panel.Controls.Add(ComboQuestion, 1, 0)
        Panel.Controls.Add(ComboAnswer, 2, 0)
        TLP.SetSize(Panel)

    End Sub
    Private Sub ButtonRemove_Enter() Handles ButtonRemove.MouseEnter
        ButtonRemove.BackgroundImage = My.Resources.Remove.ToBitmap
    End Sub
    Private Sub ButtonRemove_Leave() Handles ButtonRemove.MouseLeave
        ButtonRemove.BackgroundImage = My.Resources.RemoveGrey.ToBitmap
    End Sub
    Public Overrides Function ToString() As String
        Return Join({ComboQuestion.Name, ComboAnswer.Name}, Delimiter)
    End Function
End Class

Public Class SurroundingClass
    Public ReadOnly Property Hooked As Boolean
    Public Property Tag As Object
    Public Event Keyed(sender As Object, e As KeyPressEventArgs)
    Public Event Moused(sender As Object, e As MouseEventExtArgs)
    Private m_GlobalHook As IKeyboardMouseEvents
    Public Sub Subscribe()

        m_GlobalHook = Hook.GlobalEvents()
        AddHandler m_GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
        AddHandler m_GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
        _Hooked = True

    End Sub
    Private Sub GlobalHookKeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        RaiseEvent Keyed(Me, e)
    End Sub
    Private Sub GlobalHookMouseDownExt(ByVal sender As Object, ByVal e As MouseEventExtArgs)
        RaiseEvent Moused(Me, e)
    End Sub
    Public Sub Unsubscribe()

        If m_GlobalHook IsNot Nothing Then
            RemoveHandler m_GlobalHook.MouseDownExt, AddressOf GlobalHookMouseDownExt
            RemoveHandler m_GlobalHook.KeyPress, AddressOf GlobalHookKeyPress
            m_GlobalHook.Dispose()
        End If
        _Hooked = False

    End Sub
End Class