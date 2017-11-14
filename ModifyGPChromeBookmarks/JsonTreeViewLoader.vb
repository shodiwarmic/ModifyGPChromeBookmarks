Imports Newtonsoft.Json.Linq

Public NotInheritable Class JsonTreeViewLoader
    Private Sub New()
    End Sub

    Public Shared Sub LoadJsonToTreeView(treeView As TreeView, json As String)
        If String.IsNullOrWhiteSpace(json) Then
            Return
        End If

        treeView.Nodes.Clear()

        Try
            Dim [object] = JObject.Parse(json)
            AddObjectNodes([object], treeView)
        Catch
            Dim [object] = JArray.Parse(json)
            AddArrayNodes([object], treeView)
        End Try

        If treeView.TopNode.Text Is Nothing Or treeView.TopNode.Text Is String.Empty Then
            treeView.TopNode.Name = "BookmarkFolder"
            treeView.TopNode.Text = "Managed Bookmarks"
        End If
    End Sub

    Public Shared Sub AddObjectNodes([object] As JObject, parent As Object)
        Dim node = New TreeNode()

        For Each [property] In [object].Properties()
            Select Case [property].Name.ToString()
                Case "toplevel_name"
                    parent.Name = "BookmarkFolder"
                    parent.Text = [property].Value.ToString()
                    parent.ImageIndex = 0
                    parent.SelectedImageIndex = 0
                Case "name"
                    node.Text = [property].Value.ToString()
                Case "url"
                    node.Name = [property].Value.ToString()
                    node.ToolTipText = [property].Value.ToString()
                    node.ImageIndex = 1
                    node.SelectedImageIndex = 1
                Case "children"
                    node.Name = "BookmarkFolder"
                    AddArrayNodes(JArray.Parse([property].Value.ToString()), node)
                    node.ImageIndex = 0
                    node.SelectedImageIndex = 0
            End Select

        Next

        If node.Text.Length > 0 Then parent.Nodes.Add(node)
    End Sub

    Private Shared Sub AddArrayNodes(array As JArray, parent As Object)
        Dim node As TreeNode
        If parent.GetType Is GetType(TreeNode) Then
            node = DirectCast(parent, TreeNode)
        Else
            node = New TreeNode()
        End If

        For i = 0 To array.Count - 1
            AddTokenNodes(array(i), node)
        Next

        If parent.GetType IsNot GetType(TreeNode) Then
            parent.Nodes.Add(node)
        End If
    End Sub

    Private Shared Sub AddTokenNodes(token As JToken, parent As TreeNode)
        If TypeOf token Is JValue Then
            parent.Nodes.Add(DirectCast(token, JValue).Value)
        ElseIf TypeOf token Is JArray Then
            AddArrayNodes(DirectCast(token, JArray), parent)
        ElseIf TypeOf token Is JObject Then
            AddObjectNodes(DirectCast(token, JObject), parent)
        End If
    End Sub
End Class