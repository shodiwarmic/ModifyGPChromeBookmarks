Imports System.Windows.Forms

Public Class DialogSettings

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ComboBoxGroupPolicies_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBoxGroupPolicies.SelectedValueChanged
        Try
            If ComboBoxGroupPolicies.SelectedValue.GetType Is System.Type.GetType("System.String") Then
                TextBoxGroupPolicyUID.Text = ComboBoxGroupPolicies.SelectedValue.ToString
                ComboBoxPolicyScope.Items.Clear()
                If ModifyGPChromeBookmarks.tableGroupPolicies.Select(String.Format("PolicyGUID = '{0}'", TextBoxGroupPolicyUID.Text.ToString))(0)("UserRegFileExists") Then
                    ComboBoxPolicyScope.Items.Add("User")
                End If
                If ModifyGPChromeBookmarks.tableGroupPolicies.Select(String.Format("PolicyGUID = '{0}'", TextBoxGroupPolicyUID.Text.ToString))(0)("CompRegFileExists") Then
                    ComboBoxPolicyScope.Items.Add("Machine")
                End If
                If TextBoxGroupPolicyUID.Text = ModifyGPChromeBookmarks.strGPUID And ComboBoxPolicyScope.Items.Contains(ModifyGPChromeBookmarks.strGPScope) Then
                    ComboBoxPolicyScope.SelectedItem = ModifyGPChromeBookmarks.strGPScope
                End If
            End If
        Catch
            TextBoxGroupPolicyUID.Clear()
            ComboBoxPolicyScope.Items.Clear()
        End Try
    End Sub
End Class
