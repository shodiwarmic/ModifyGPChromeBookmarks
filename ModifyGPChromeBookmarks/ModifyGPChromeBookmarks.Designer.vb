<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ModifyGPChromeBookmarks
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ModifyGPChromeBookmarks))
        Me.ButtonApplyChanges = New System.Windows.Forms.Button()
        Me.MainMenuStrip = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TextBoxPolicyGUID = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TextBoxPolicyName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxPolicyScope = New System.Windows.Forms.TextBox()
        Me.TreeViewBookmarks = New System.Windows.Forms.TreeView()
        Me.ImageListTreeView = New System.Windows.Forms.ImageList(Me.components)
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBoxLabel = New System.Windows.Forms.TextBox()
        Me.TextBoxURL = New System.Windows.Forms.TextBox()
        Me.ButtonAddFolder = New System.Windows.Forms.Button()
        Me.ButtonAddShortcut = New System.Windows.Forms.Button()
        Me.ButtonDeleteNode = New System.Windows.Forms.Button()
        Me.MainMenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonApplyChanges
        '
        Me.ButtonApplyChanges.Location = New System.Drawing.Point(242, 468)
        Me.ButtonApplyChanges.Name = "ButtonApplyChanges"
        Me.ButtonApplyChanges.Size = New System.Drawing.Size(148, 23)
        Me.ButtonApplyChanges.TabIndex = 15
        Me.ButtonApplyChanges.Text = "Apply Changes"
        Me.ButtonApplyChanges.UseVisualStyleBackColor = True
        '
        'MainMenuStrip
        '
        Me.MainMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MainMenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MainMenuStrip.Name = "MainMenuStrip"
        Me.MainMenuStrip.Size = New System.Drawing.Size(402, 24)
        Me.MainMenuStrip.TabIndex = 0
        Me.MainMenuStrip.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.QuitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'QuitToolStripMenuItem
        '
        Me.QuitToolStripMenuItem.Name = "QuitToolStripMenuItem"
        Me.QuitToolStripMenuItem.Size = New System.Drawing.Size(97, 22)
        Me.QuitToolStripMenuItem.Text = "Quit"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SettingsToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(47, 20)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'SettingsToolStripMenuItem
        '
        Me.SettingsToolStripMenuItem.Name = "SettingsToolStripMenuItem"
        Me.SettingsToolStripMenuItem.Size = New System.Drawing.Size(116, 22)
        Me.SettingsToolStripMenuItem.Text = "Settings"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'TextBoxPolicyGUID
        '
        Me.TextBoxPolicyGUID.Enabled = False
        Me.TextBoxPolicyGUID.Location = New System.Drawing.Point(84, 53)
        Me.TextBoxPolicyGUID.Name = "TextBoxPolicyGUID"
        Me.TextBoxPolicyGUID.Size = New System.Drawing.Size(306, 20)
        Me.TextBoxPolicyGUID.TabIndex = 6
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(13, 56)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(65, 13)
        Me.Label9.TabIndex = 5
        Me.Label9.Text = "Policy GUID"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 30)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(66, 13)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Policy Name"
        '
        'TextBoxPolicyName
        '
        Me.TextBoxPolicyName.Enabled = False
        Me.TextBoxPolicyName.Location = New System.Drawing.Point(84, 27)
        Me.TextBoxPolicyName.Name = "TextBoxPolicyName"
        Me.TextBoxPolicyName.Size = New System.Drawing.Size(161, 20)
        Me.TextBoxPolicyName.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(251, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Policy Scope"
        '
        'TextBoxPolicyScope
        '
        Me.TextBoxPolicyScope.Enabled = False
        Me.TextBoxPolicyScope.Location = New System.Drawing.Point(326, 27)
        Me.TextBoxPolicyScope.Name = "TextBoxPolicyScope"
        Me.TextBoxPolicyScope.Size = New System.Drawing.Size(64, 20)
        Me.TextBoxPolicyScope.TabIndex = 4
        '
        'TreeViewBookmarks
        '
        Me.TreeViewBookmarks.AllowDrop = True
        Me.TreeViewBookmarks.HideSelection = False
        Me.TreeViewBookmarks.ImageIndex = 0
        Me.TreeViewBookmarks.ImageList = Me.ImageListTreeView
        Me.TreeViewBookmarks.Location = New System.Drawing.Point(12, 79)
        Me.TreeViewBookmarks.Name = "TreeViewBookmarks"
        Me.TreeViewBookmarks.SelectedImageIndex = 0
        Me.TreeViewBookmarks.ShowNodeToolTips = True
        Me.TreeViewBookmarks.Size = New System.Drawing.Size(378, 302)
        Me.TreeViewBookmarks.TabIndex = 7
        '
        'ImageListTreeView
        '
        Me.ImageListTreeView.ImageStream = CType(resources.GetObject("ImageListTreeView.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListTreeView.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListTreeView.Images.SetKeyName(0, "folder.png")
        Me.ImageListTreeView.Images.SetKeyName(1, "node.png")
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 445)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "URL: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 419)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Label: "
        '
        'TextBoxLabel
        '
        Me.TextBoxLabel.Enabled = False
        Me.TextBoxLabel.Location = New System.Drawing.Point(57, 416)
        Me.TextBoxLabel.Name = "TextBoxLabel"
        Me.TextBoxLabel.Size = New System.Drawing.Size(333, 20)
        Me.TextBoxLabel.TabIndex = 12
        '
        'TextBoxURL
        '
        Me.TextBoxURL.Enabled = False
        Me.TextBoxURL.Location = New System.Drawing.Point(57, 442)
        Me.TextBoxURL.Name = "TextBoxURL"
        Me.TextBoxURL.Size = New System.Drawing.Size(333, 20)
        Me.TextBoxURL.TabIndex = 14
        '
        'ButtonAddFolder
        '
        Me.ButtonAddFolder.Location = New System.Drawing.Point(12, 387)
        Me.ButtonAddFolder.Name = "ButtonAddFolder"
        Me.ButtonAddFolder.Size = New System.Drawing.Size(122, 23)
        Me.ButtonAddFolder.TabIndex = 8
        Me.ButtonAddFolder.Text = "Add Folder"
        Me.ButtonAddFolder.UseVisualStyleBackColor = True
        '
        'ButtonAddShortcut
        '
        Me.ButtonAddShortcut.Location = New System.Drawing.Point(140, 387)
        Me.ButtonAddShortcut.Name = "ButtonAddShortcut"
        Me.ButtonAddShortcut.Size = New System.Drawing.Size(122, 23)
        Me.ButtonAddShortcut.TabIndex = 9
        Me.ButtonAddShortcut.Text = "Add Shortcut"
        Me.ButtonAddShortcut.UseVisualStyleBackColor = True
        '
        'ButtonDeleteNode
        '
        Me.ButtonDeleteNode.Location = New System.Drawing.Point(268, 387)
        Me.ButtonDeleteNode.Name = "ButtonDeleteNode"
        Me.ButtonDeleteNode.Size = New System.Drawing.Size(122, 23)
        Me.ButtonDeleteNode.TabIndex = 10
        Me.ButtonDeleteNode.Text = "Delete Item"
        Me.ButtonDeleteNode.UseVisualStyleBackColor = True
        '
        'ModifyGPChromeBookmarks
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 503)
        Me.Controls.Add(Me.ButtonDeleteNode)
        Me.Controls.Add(Me.ButtonAddShortcut)
        Me.Controls.Add(Me.ButtonAddFolder)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBoxLabel)
        Me.Controls.Add(Me.TextBoxURL)
        Me.Controls.Add(Me.TreeViewBookmarks)
        Me.Controls.Add(Me.TextBoxPolicyScope)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBoxPolicyGUID)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.TextBoxPolicyName)
        Me.Controls.Add(Me.ButtonApplyChanges)
        Me.Controls.Add(Me.MainMenuStrip)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "ModifyGPChromeBookmarks"
        Me.Text = "Modify GP Chrome Bookmarks"
        Me.MainMenuStrip.ResumeLayout(False)
        Me.MainMenuStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ButtonApplyChanges As Button
    Friend WithEvents MainMenuStrip As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents QuitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TextBoxPolicyGUID As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents TextBoxPolicyName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxPolicyScope As TextBox
    Friend WithEvents TreeViewBookmarks As TreeView
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBoxLabel As TextBox
    Friend WithEvents TextBoxURL As TextBox
    Friend WithEvents ImageListTreeView As ImageList
    Friend WithEvents ButtonAddFolder As Button
    Friend WithEvents ButtonAddShortcut As Button
    Friend WithEvents ButtonDeleteNode As Button
End Class
