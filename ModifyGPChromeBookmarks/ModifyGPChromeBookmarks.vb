Imports MadMilkman.Ini
Imports ModifyGPChromeBookmarks.RegistryPolicyClass
Imports System.DirectoryServices
Imports System.IO
Imports Microsoft.Win32
Imports ModifyGPChromeBookmarks.JsonTreeViewLoader
Imports System.Text
Imports System.DirectoryServices.ActiveDirectory

Public Class ModifyGPChromeBookmarks
    Private NodeMap As String
    Private Const MAPSIZE As Integer = 128
    Private NewNodeMap As New StringBuilder(MAPSIZE)

    Dim objDomain As Domain = Domain.GetComputerDomain()
    Dim strDomainName As String = objDomain.Name
    Dim strGPPath, strValue As String
    Protected Friend strGPUID, strGPScope As String

    Dim regUserSettings As RegistryKey

    Protected Friend tableGroupPolicies As New DataTable
    Dim objPolFile As New PolFile

    Private Sub MakeDataTables()
        tableGroupPolicies.Columns.Add("PolicyGUID")
        tableGroupPolicies.Columns.Add("PolicyName")
        tableGroupPolicies.Columns.Add("UserRegFileExists", Type.GetType("System.Boolean"))
        tableGroupPolicies.Columns.Add("CompRegFileExists", Type.GetType("System.Boolean"))
    End Sub

    Private Sub RefreshPolicyList()
        ' Remove any old polices from the table
        tableGroupPolicies.Clear()

        ' Populate the group policy table with all group policies in the domain
        Using searcher = New DirectorySearcher(objDomain.GetDirectoryEntry, "(&(objectClass=groupPolicyContainer))", {"displayName", "name"}, SearchScope.Subtree)
            Dim results As SearchResultCollection = searcher.FindAll
            For Each result As SearchResult In results
                Dim directoryEntry As DirectoryEntry = result.GetDirectoryEntry
                Dim row As DataRow
                Try
                    row = tableGroupPolicies.Select(String.Format("PolicyGUID = '{0}'", directoryEntry.Properties("name").Value))(0)
                Catch
                    row = tableGroupPolicies.NewRow
                    Dim strTempGPPath As String = "\\" + strDomainName + "\SYSVOL\" + strDomainName + "\Policies\" + directoryEntry.Properties("name").Value
                    row("PolicyGUID") = directoryEntry.Properties("name").Value
                    row("PolicyName") = directoryEntry.Properties("displayName").Value
                    ' Check if the needed group policy settings files already exist
                    row("UserRegFileExists") = File.Exists(strTempGPPath + "\User\registry.pol")
                    row("CompRegFileExists") = File.Exists(strTempGPPath + "\Machine\registry.pol")
                    tableGroupPolicies.Rows.Add(row)
                End Try
            Next
        End Using

        tableGroupPolicies.DefaultView.RowFilter = "UserRegFileExists = TRUE OR CompRegFileExists = TRUE"
    End Sub

    Private Sub LaunchSettings()
        ' Check if policies have changed
        RefreshPolicyList()

        ' Set the settings dropdown list to show the found policy names, but return the guid of the selected policy
        DialogSettings.ComboBoxGroupPolicies.DataSource = tableGroupPolicies
        DialogSettings.ComboBoxGroupPolicies.DisplayMember = "PolicyName"
        DialogSettings.ComboBoxGroupPolicies.ValueMember = "PolicyGUID"

        ' Preselect the current policy, if there is one
        If strGPUID <> Nothing And strGPUID <> String.Empty Then DialogSettings.ComboBoxGroupPolicies.SelectedValue = strGPUID

        If DialogSettings.ShowDialog() = DialogResult.OK Then

            strGPUID = DialogSettings.ComboBoxGroupPolicies.SelectedValue
            strGPScope = DialogSettings.ComboBoxPolicyScope.Text
            strGPPath = "\\" + strDomainName + "\SYSVOL\" + strDomainName + "\Policies\" + strGPUID

            ' Save this to the registry for the next time the program is openned
            regUserSettings.SetValue("GPUID", strGPUID)
            regUserSettings.SetValue("GPScope", strGPScope)

            ' Set the Policy Name and GUID text boxes on the main form
            TextBoxPolicyName.Text = tableGroupPolicies.Select(String.Format("PolicyGUID = '{0}'", strGPUID))(0)("PolicyName")
            TextBoxPolicyGUID.Text = strGPUID
            TextBoxPolicyScope.Text = strGPScope

            LoadBookmarkList()
        End If

    End Sub

    Private Sub LoadBookmarkList()
        TreeViewBookmarks.Nodes.Clear()

        If (objPolFile.LoadFile(strGPPath + "\" + strGPScope + "\registry.pol") <> PF_SUCCESS) Then
            objPolFile = New PolFile
        End If

        If (objPolFile.GetValue("Software\Policies\Google\Chrome\ManagedBookmarks", strValue, REG_DWORD) = PF_SUCCESS) Then
            If (strValue.Contains("name") And strValue.Contains("url")) Then
                LoadJsonToTreeView(TreeViewBookmarks, strValue)
            End If
        End If

        If TreeViewBookmarks.Nodes.Count < 1 Then
            TreeViewBookmarks.Nodes.Add(New TreeNode("Managed Bookmarks", 0, 0))
            TreeViewBookmarks.TopNode.Name = "BookmarkFolder"
        End If

    End Sub

    Private Sub ModifyGPChromeBookmarks_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Create the data tables
        MakeDataTables()

        ' Load list of all policies in domain
        RefreshPolicyList()

        ' Open the registry settings for the application
        regUserSettings = Registry.CurrentUser.OpenSubKey("Software\GroupPolicyChromeBookmarks", True)

        If regUserSettings Is Nothing Then
            ' Create the registry settings if they do not exist
            regUserSettings = Registry.CurrentUser.CreateSubKey("Software\GroupPolicyChromeBookmarks")
        ElseIf strGPUID Is Nothing Or strGPUID = String.Empty Or strGPScope Is Nothing Or strGPScope = String.Empty Then
            Try
                ' Attempt to set the current policy guid and scope from the registry
                strGPUID = regUserSettings.GetValue("GPUID").ToString()
                strGPScope = regUserSettings.GetValue("GPScope").ToString()
                strGPPath = "\\" + strDomainName + "\SYSVOL\" + strDomainName + "\Policies\" + strGPUID
                TextBoxPolicyName.Text = tableGroupPolicies.Select(String.Format("PolicyGUID = '{0}'", strGPUID))(0)("PolicyName")
                TextBoxPolicyGUID.Text = strGPUID
                TextBoxPolicyScope.Text = strGPScope
            Catch
            End Try
        Else
            ' Set the registry setting to the current policy UID
            regUserSettings.SetValue("GPUID", strGPUID)
            regUserSettings.SetValue("GPScope", strGPScope)
        End If

        Do While strGPUID Is Nothing Or strGPUID = String.Empty Or strGPScope Is Nothing Or strGPScope = String.Empty
            ' Force user to select a group policy
            LaunchSettings()
            If strGPUID Is Nothing Or strGPUID = String.Empty Or strGPScope Is Nothing Or strGPScope = String.Empty Then
                MessageBox.Show("You must select a group policy and scope on your first launch.", "First Launch", MessageBoxButtons.OK)
            End If
        Loop

        LoadBookmarkList()

    End Sub

    Private Sub ButtonApplyChanges_Click(sender As Object, e As EventArgs) Handles ButtonApplyChanges.Click
        strValue = "[{""toplevel_name"": """ + TreeViewBookmarks.TopNode.Text + """}, "
        For Each node As TreeNode In TreeViewBookmarks.Nodes
            strValue += EvaluateChildNodes(node)
        Next
        strValue += "]"

        Dim intRet As Integer

        If (strValue = "[{""toplevel_name"": """ + TreeViewBookmarks.TopNode.Text + """}, ]") Then
            intRet = objPolFile.DeleteValue("Software\Policies\Google\Chrome\ManagedBookmarks")
        Else
            intRet = objPolFile.SetValue("Software\Policies\Google\Chrome", "ManagedBookmarks", REG_SZ, strValue)
        End If

        If (intRet = PF_SUCCESS) Then
            Dim iniFile As New IniFile()
            iniFile.Load(strGPPath + "\GPT.ini")
            Dim Version = iniFile.Sections.Item("General").Keys.Item("Version").Value
            iniFile.Sections.Item("General").Keys.Item("Version").Value = Version + 1
            iniFile.Save(strGPPath + "\GPT.ini")
            objPolFile.Save(strGPPath + "\" + strGPScope + "\registry.pol", True)
            MsgBox("Changes saved to policy file.", MsgBoxStyle.Information, "Saved!")
        Else
            MsgBox("Unable to save policy file.", MsgBoxStyle.Exclamation, "Error!")
        End If


    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox.Visible = True
        AboutBox.Activate()
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        End
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        LaunchSettings()
    End Sub

    Private Function EvaluateChildNodes(node As TreeNode) As String
        Dim str As String = ""
        For Each child As TreeNode In node.Nodes
            If str.Count > 1 Then str += ", "
            str += "{""name"": """ + child.Text + """, "
            If child.Name = "BookmarkFolder" Then
                str += """children"": [" + EvaluateChildNodes(child) + "]}"
            Else
                str += """url"": """ + child.Name + """}"
            End If
        Next
        Return str
    End Function

    Private Sub TreeViewBookmarks_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeViewBookmarks.AfterSelect
        ButtonDeleteNode.Enabled = True
        TextBoxLabel.Text = TreeViewBookmarks.SelectedNode.Text
        TextBoxLabel.Enabled = True
        If TreeViewBookmarks.SelectedNode.Name = "BookmarkFolder" Then
            TextBoxURL.Clear()
            TextBoxURL.Enabled = False
        Else
            TextBoxURL.Text = TreeViewBookmarks.SelectedNode.Name
            TextBoxURL.Enabled = True
        End If
    End Sub

    Private Sub TextBoxURL_TextChanged(sender As Object, e As EventArgs) Handles TextBoxURL.TextChanged
        If Not TreeViewBookmarks.SelectedNode.Name = "BookmarkFolder" Then
            TreeViewBookmarks.SelectedNode.Name = TextBoxURL.Text
            TreeViewBookmarks.SelectedNode.ToolTipText = TextBoxURL.Text
        End If
    End Sub

    Private Sub TextBoxLabel_TextChanged(sender As Object, e As EventArgs) Handles TextBoxLabel.TextChanged
        TreeViewBookmarks.SelectedNode.Text = TextBoxLabel.Text
    End Sub

    Private Sub ButtonAddFolder_Click(sender As Object, e As EventArgs) Handles ButtonAddFolder.Click
        Dim node As New TreeNode("Shortcut Folder", 0, 0)
        node.Name = "BookmarkFolder"
        TreeViewBookmarks.TopNode.Nodes.Add(node)
    End Sub

    Private Sub ButtonAddShortcut_Click(sender As Object, e As EventArgs) Handles ButtonAddShortcut.Click
        TreeViewBookmarks.TopNode.Nodes.Add(New TreeNode("New Shortcut", 1, 1))
    End Sub

    Private Sub ButtonDeleteNode_Click(sender As Object, e As EventArgs) Handles ButtonDeleteNode.Click
        If TreeViewBookmarks.SelectedNode Is TreeViewBookmarks.TopNode Then
            MessageBox.Show("You can not delete the top node.", "Unable to delete", MessageBoxButtons.OK)
        Else
            Dim prompt As String
            If TreeViewBookmarks.SelectedNode.Name = "BookmarkFolder" Then
                prompt = "Are you sure you wish to delete this folder?"
                If TreeViewBookmarks.SelectedNode.Nodes.Count > 0 Then
                    prompt += vbCrLf + "All contained shortcuts and folders will be deleted."
                End If
            Else
                prompt = "Are you sure you wish to delete this shortcut?"
            End If
            If MessageBox.Show(prompt, "Are you sure?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                TreeViewBookmarks.Nodes.Remove(TreeViewBookmarks.SelectedNode)
            End If
        End If
    End Sub




    Private Sub TreeViewBookmarks_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles TreeViewBookmarks.MouseDown
        Dim node As TreeNode = TreeViewBookmarks.GetNodeAt(e.X, e.Y)
        If node IsNot Nothing Then TreeViewBookmarks.SelectedNode = node
    End Sub
    Private Sub TreeViewBookmarks_ItemDrag(ByVal sender As Object, ByVal e As ItemDragEventArgs) Handles TreeViewBookmarks.ItemDrag
        DoDragDrop(e.Item, DragDropEffects.Move)
    End Sub
    Private Sub TreeViewBookmarks_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles TreeViewBookmarks.DragEnter
        e.Effect = DragDropEffects.Move
    End Sub
    Private Sub TreeViewBookmarks_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles TreeViewBookmarks.DragDrop
        If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", False) AndAlso Not (NodeMap = "") Then
            Dim MovingNode As TreeNode = CType(e.Data.GetData("System.Windows.Forms.TreeNode"), TreeNode)
            Dim NodeIndexes As String() = NodeMap.Split("|"c)
            Dim InsertCollection As TreeNodeCollection = TreeViewBookmarks.Nodes
            Dim i As Integer = 0
            While i < NodeIndexes.Length - 1
                InsertCollection = InsertCollection(Int32.Parse(NodeIndexes(i))).Nodes
                System.Math.Min(System.Threading.Interlocked.Increment(i), i - 1)
            End While
            If Not (InsertCollection Is Nothing) Then
                InsertCollection.Insert(Int32.Parse(NodeIndexes(NodeIndexes.Length - 1)), CType(MovingNode.Clone, TreeNode))
                TreeViewBookmarks.SelectedNode = InsertCollection(Int32.Parse(NodeIndexes(NodeIndexes.Length - 1)))
                MovingNode.Remove()
            End If
        End If
    End Sub
    Private Sub TreeViewBookmarks_DragOver(ByVal sender As Object, ByVal e As DragEventArgs) Handles TreeViewBookmarks.DragOver
        Dim NodeOver As TreeNode = TreeViewBookmarks.GetNodeAt(TreeViewBookmarks.PointToClient(Cursor.Position))
        Dim NodeMoving As TreeNode = CType(e.Data.GetData("System.Windows.Forms.TreeNode"), TreeNode)
        If Not (NodeOver Is Nothing) AndAlso (Not (NodeOver Is NodeMoving) OrElse (Not (NodeOver.Parent Is Nothing) AndAlso NodeOver.Index = (NodeOver.Parent.Nodes.Count - 1))) Then
            Dim OffsetY As Integer = TreeViewBookmarks.PointToClient(Cursor.Position).Y - NodeOver.Bounds.Top
            Dim NodeOverImageWidth As Integer = TreeViewBookmarks.ImageList.Images(NodeOver.ImageIndex).Size.Width + 8
            Dim g As Graphics = TreeViewBookmarks.CreateGraphics
            If NodeOver.ImageIndex = 1 Then
                If OffsetY < (NodeOver.Bounds.Height / 2) Then
                    Dim tnParadox As TreeNode = NodeOver
                    While Not (tnParadox.Parent Is Nothing)
                        If tnParadox.Parent Is NodeMoving Then
                            NodeMap = ""
                            Return
                        End If
                        tnParadox = tnParadox.Parent
                    End While
                    SetNewNodeMap(NodeOver, False)  ' ERROR!  Not sure why
                    If SetMapsEqual() = True Then
                        Return
                    End If
                    Refresh()
                    DrawLeafTopPlaceholders(NodeOver)
                Else
                    Dim tnParadox As TreeNode = NodeOver
                    While Not (tnParadox.Parent Is Nothing)
                        If tnParadox.Parent Is NodeMoving Then
                            NodeMap = ""
                            Return
                        End If
                        tnParadox = tnParadox.Parent
                    End While
                    Dim ParentDragDrop As TreeNode = Nothing
                    If Not (NodeOver.Parent Is Nothing) AndAlso NodeOver.Index = (NodeOver.Parent.Nodes.Count - 1) Then
                        Dim XPos As Integer = TreeViewBookmarks.PointToClient(Cursor.Position).X
                        If XPos < NodeOver.Bounds.Left Then
                            ParentDragDrop = NodeOver.Parent
                            If XPos < (ParentDragDrop.Bounds.Left - TreeViewBookmarks.ImageList.Images(ParentDragDrop.ImageIndex).Size.Width) Then
                                If Not (ParentDragDrop.Parent Is Nothing) Then
                                    ParentDragDrop = ParentDragDrop.Parent
                                End If
                            End If
                        End If
                    End If
                    SetNewNodeMap(Microsoft.VisualBasic.IIf(Not (ParentDragDrop Is Nothing), ParentDragDrop, NodeOver), True)
                    If SetMapsEqual() = True Then
                        Return
                    End If
                    Refresh()
                    DrawLeafBottomPlaceholders(NodeOver, ParentDragDrop)
                End If
            Else
                If OffsetY < (NodeOver.Bounds.Height / 3) Then
                    Dim tnParadox As TreeNode = NodeOver
                    While Not (tnParadox.Parent Is Nothing)
                        If tnParadox.Parent Is NodeMoving Then
                            NodeMap = ""
                            Return
                        End If
                        tnParadox = tnParadox.Parent
                    End While
                    SetNewNodeMap(NodeOver, False)
                    If SetMapsEqual() = True Then
                        Return
                    End If
                    Refresh()
                    DrawFolderTopPlaceholders(NodeOver)
                Else
                    If (Not (NodeOver.Parent Is Nothing) AndAlso NodeOver.Index = 0) AndAlso (OffsetY > (NodeOver.Bounds.Height - (NodeOver.Bounds.Height / 3))) Then
                        Dim tnParadox As TreeNode = NodeOver
                        While Not (tnParadox.Parent Is Nothing)
                            If tnParadox.Parent Is NodeMoving Then
                                NodeMap = ""
                                Return
                            End If
                            tnParadox = tnParadox.Parent
                        End While
                        SetNewNodeMap(NodeOver, True)
                        If SetMapsEqual() = True Then
                            Return
                        End If
                        Refresh()
                        DrawFolderTopPlaceholders(NodeOver)
                    Else
                        If NodeOver.Nodes.Count > 0 Then
                            NodeOver.Expand()
                        Else
                            If NodeMoving Is NodeOver Then
                                Return
                            End If
                            Dim tnParadox As TreeNode = NodeOver
                            While Not (tnParadox.Parent Is Nothing)
                                If tnParadox.Parent Is NodeMoving Then
                                    NodeMap = ""
                                    Return
                                End If
                                tnParadox = tnParadox.Parent
                            End While
                            SetNewNodeMap(NodeOver, False)
                            NewNodeMap = NewNodeMap.Insert(NewNodeMap.Length, "|0")
                            If SetMapsEqual() = True Then
                                Return
                            End If
                            Refresh()
                            DrawAddToFolderPlaceholder(NodeOver)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub DrawLeafTopPlaceholders(ByVal NodeOver As TreeNode)
        Dim g As Graphics = TreeViewBookmarks.CreateGraphics
        Dim NodeOverImageWidth As Integer = TreeViewBookmarks.ImageList.Images(NodeOver.ImageIndex).Size.Width + 8
        Dim LeftPos As Integer = NodeOver.Bounds.Left - NodeOverImageWidth
        Dim RightPos As Integer = TreeViewBookmarks.Width - 4
        Dim LeftTriangle As Point() = New Point(4) {New Point(LeftPos, NodeOver.Bounds.Top - 4), New Point(LeftPos, NodeOver.Bounds.Top + 4), New Point(LeftPos + 4, NodeOver.Bounds.Y), New Point(LeftPos + 4, NodeOver.Bounds.Top - 1), New Point(LeftPos, NodeOver.Bounds.Top - 5)}
        Dim RightTriangle As Point() = New Point(4) {New Point(RightPos, NodeOver.Bounds.Top - 4), New Point(RightPos, NodeOver.Bounds.Top + 4), New Point(RightPos - 4, NodeOver.Bounds.Y), New Point(RightPos - 4, NodeOver.Bounds.Top - 1), New Point(RightPos, NodeOver.Bounds.Top - 5)}
        g.FillPolygon(System.Drawing.Brushes.Black, LeftTriangle)
        g.FillPolygon(System.Drawing.Brushes.Black, RightTriangle)
        g.DrawLine(New System.Drawing.Pen(Color.Black, 2), New Point(LeftPos, NodeOver.Bounds.Top), New Point(RightPos, NodeOver.Bounds.Top))
    End Sub
    Private Sub DrawLeafBottomPlaceholders(ByVal NodeOver As TreeNode, ByVal ParentDragDrop As TreeNode)
        Dim g As Graphics = TreeViewBookmarks.CreateGraphics
        Dim NodeOverImageWidth As Integer = TreeViewBookmarks.ImageList.Images(NodeOver.ImageIndex).Size.Width + 8
        Dim LeftPos As Integer
        Dim RightPos As Integer
        If Not (ParentDragDrop Is Nothing) Then
            LeftPos = ParentDragDrop.Bounds.Left - (TreeViewBookmarks.ImageList.Images(ParentDragDrop.ImageIndex).Size.Width + 8)
        Else
            LeftPos = NodeOver.Bounds.Left - NodeOverImageWidth
        End If
        RightPos = TreeViewBookmarks.Width - 4
        Dim LeftTriangle As Point() = New Point(4) {New Point(LeftPos, NodeOver.Bounds.Bottom - 4), New Point(LeftPos, NodeOver.Bounds.Bottom + 4), New Point(LeftPos + 4, NodeOver.Bounds.Bottom), New Point(LeftPos + 4, NodeOver.Bounds.Bottom - 1), New Point(LeftPos, NodeOver.Bounds.Bottom - 5)}
        Dim RightTriangle As Point() = New Point(4) {New Point(RightPos, NodeOver.Bounds.Bottom - 4), New Point(RightPos, NodeOver.Bounds.Bottom + 4), New Point(RightPos - 4, NodeOver.Bounds.Bottom), New Point(RightPos - 4, NodeOver.Bounds.Bottom - 1), New Point(RightPos, NodeOver.Bounds.Bottom - 5)}
        g.FillPolygon(System.Drawing.Brushes.Black, LeftTriangle)
        g.FillPolygon(System.Drawing.Brushes.Black, RightTriangle)
        g.DrawLine(New System.Drawing.Pen(Color.Black, 2), New Point(LeftPos, NodeOver.Bounds.Bottom), New Point(RightPos, NodeOver.Bounds.Bottom))
    End Sub
    Private Sub DrawFolderTopPlaceholders(ByVal NodeOver As TreeNode)
        Dim g As Graphics = TreeViewBookmarks.CreateGraphics
        Dim NodeOverImageWidth As Integer = TreeViewBookmarks.ImageList.Images(NodeOver.ImageIndex).Size.Width + 8
        Dim LeftPos As Integer
        Dim RightPos As Integer
        LeftPos = NodeOver.Bounds.Left - NodeOverImageWidth
        RightPos = TreeViewBookmarks.Width - 4
        Dim LeftTriangle As Point() = New Point(4) {New Point(LeftPos, NodeOver.Bounds.Top - 4), New Point(LeftPos, NodeOver.Bounds.Top + 4), New Point(LeftPos + 4, NodeOver.Bounds.Y), New Point(LeftPos + 4, NodeOver.Bounds.Top - 1), New Point(LeftPos, NodeOver.Bounds.Top - 5)}
        Dim RightTriangle As Point() = New Point(4) {New Point(RightPos, NodeOver.Bounds.Top - 4), New Point(RightPos, NodeOver.Bounds.Top + 4), New Point(RightPos - 4, NodeOver.Bounds.Y), New Point(RightPos - 4, NodeOver.Bounds.Top - 1), New Point(RightPos, NodeOver.Bounds.Top - 5)}
        g.FillPolygon(System.Drawing.Brushes.Black, LeftTriangle)
        g.FillPolygon(System.Drawing.Brushes.Black, RightTriangle)
        g.DrawLine(New System.Drawing.Pen(Color.Black, 2), New Point(LeftPos, NodeOver.Bounds.Top), New Point(RightPos, NodeOver.Bounds.Top))
    End Sub
    Private Sub DrawAddToFolderPlaceholder(ByVal NodeOver As TreeNode)
        Dim g As Graphics = TreeViewBookmarks.CreateGraphics
        Dim RightPos As Integer = NodeOver.Bounds.Right + 6
        Dim RightTriangle As Point() = New Point(4) {New Point(RightPos, NodeOver.Bounds.Y + (NodeOver.Bounds.Height / 2) + 4), New Point(RightPos, NodeOver.Bounds.Y + (NodeOver.Bounds.Height / 2) + 4), New Point(RightPos - 4, NodeOver.Bounds.Y + (NodeOver.Bounds.Height / 2)), New Point(RightPos - 4, NodeOver.Bounds.Y + (NodeOver.Bounds.Height / 2) - 1), New Point(RightPos, NodeOver.Bounds.Y + (NodeOver.Bounds.Height / 2) - 5)}
        Refresh()
        g.FillPolygon(System.Drawing.Brushes.Black, RightTriangle)
    End Sub
    Private Sub SetNewNodeMap(ByVal tnNode As TreeNode, ByVal boolBelowNode As Boolean)
        NewNodeMap.Length = 0
        If boolBelowNode Then
            NewNodeMap.Insert(0, CType(tnNode.Index, Integer) + 1)
        Else
            NewNodeMap.Insert(0, CType(tnNode.Index, Integer))
        End If
        Dim tnCurNode As TreeNode = tnNode
        While Not (tnCurNode.Parent Is Nothing)
            tnCurNode = tnCurNode.Parent
            If NewNodeMap.Length = 0 AndAlso boolBelowNode = True Then
                NewNodeMap.Insert(0, (tnCurNode.Index + 1).ToString() + "|")
            Else
                NewNodeMap.Insert(0, tnCurNode.Index.ToString() + "|")
            End If
        End While
    End Sub
    Private Function SetMapsEqual() As Boolean
        If NewNodeMap.ToString = NodeMap Then
            Return True
        Else
            NodeMap = NewNodeMap.ToString
            Return False
        End If
    End Function



End Class
