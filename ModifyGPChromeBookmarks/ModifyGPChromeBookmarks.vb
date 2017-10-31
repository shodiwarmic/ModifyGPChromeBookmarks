Imports MadMilkman.Ini
Imports Newtonsoft.Json
Imports ModifyGPChromeBookmarks.RegistryPolicyClass
Imports System.DirectoryServices
Imports System.IO
Imports Microsoft.Win32

Public Class ModifyGPChromeBookmarks

    Dim objDomain As ActiveDirectory.Domain = ActiveDirectory.Domain.GetComputerDomain()
    Dim strDomainName As String = objDomain.Name
    Dim strGPPath, strValue As String
    Protected Friend strGPUID, strGPScope As String

    Dim regUserSettings As RegistryKey

    Dim tableSites As New DataTable()
    Protected Friend tableGroupPolicies As New DataTable
    Dim objPolFile As New PolFile

    Private Sub MakeDataTables()
        tableSites.Columns.Add("name")
        tableSites.Columns.Add("url")
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
        tableSites.Clear()

        If (objPolFile.LoadFile(strGPPath + "\" + strGPScope + "\registry.pol") <> PF_SUCCESS) Then
            objPolFile = New PolFile
        End If

        If (objPolFile.GetValue("Software\Policies\Google\Chrome\ManagedBookmarks", strValue, REG_DWORD) = PF_SUCCESS) Then
            If (strValue.Contains("name") And strValue.Contains("url")) Then
                tableSites = JsonConvert.DeserializeObject(Of DataTable)(strValue)
            End If
        End If

        gvBookmarks.DataSource = tableSites

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

        gvBookmarks.Columns.Item(0).Width = 150
        gvBookmarks.Columns.Item(1).Width = 480

    End Sub

    Private Sub btnApplyChanges_Click(sender As Object, e As EventArgs) Handles btnApplyChanges.Click
        strValue = JsonConvert.SerializeObject(tableSites)
        Dim intRet As Integer

        If (strValue = "[]") Then
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
End Class
