VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'DAG (directed acyclic graph) algorithm support class

'... TODO:
'topological sorting: easy
'strongly connected components: Tarjan's Algorithm

Private Type typeGraphNode
 obj As IASTNode 'IUnknown
 nDegreeIn As Long
 nDegreeOut As Long
 nOut() As Long '1-based index array
End Type

Private m_tNodes() As typeGraphNode '1-based
Private m_nNodeCount As Long

Private m_nSortedNode() As Long

Private m_objNode As New Collection

'L <-- Empty list that will contain the sorted elements
'S <-- Set of all nodes with no incoming edges
'while S is non-empty do
'    remove a node n from S
'    insert n into L
'    for each node m with an edge e from n to m do
'        remove edge e from the graph
'        if m has no other incoming edges then
'            insert m into S
'if graph has edges then
'    output error message (graph has at least one cycle)
'else
'    output message (proposed topologically sorted order: L)
Friend Function RunTopologicalSort() As Boolean
Dim i As Long, j As Long
Dim n As Long, m As Long
Dim tmp As Long
Dim nEmptyNodes() As Long 'stack, 1-based
Dim nEmptyNodeCount As Long
Dim nSortedNodeCount As Long
'///
If m_nNodeCount <= 0 Then
 RunTopologicalSort = True
 Exit Function
End If
'///
ReDim m_nSortedNode(1 To m_nNodeCount)
ReDim nEmptyNodes(1 To m_nNodeCount)
'///
For i = 1 To m_nNodeCount
 Debug.Assert m_tNodes(i).nDegreeIn >= 0
 If m_tNodes(i).nDegreeIn = 0 Then
  nEmptyNodeCount = nEmptyNodeCount + 1
  nEmptyNodes(nEmptyNodeCount) = i
 End If
Next i
'///
Do While nEmptyNodeCount > 0
 n = nEmptyNodes(nEmptyNodeCount)
 nEmptyNodeCount = nEmptyNodeCount - 1
 '///
 nSortedNodeCount = nSortedNodeCount + 1
 m_nSortedNode(nSortedNodeCount) = n
 '///
 For j = 1 To m_tNodes(n).nDegreeOut
  m = m_tNodes(n).nOut(j)
  tmp = m_tNodes(m).nDegreeIn - 1
  m_tNodes(m).nDegreeIn = tmp
  Debug.Assert tmp >= 0
  If tmp = 0 Then
   nEmptyNodeCount = nEmptyNodeCount + 1
   nEmptyNodes(nEmptyNodeCount) = m
  End If
 Next j
Loop
'///
RunTopologicalSort = nSortedNodeCount = m_nNodeCount
End Function

Friend Sub Clear()
Erase m_tNodes, m_nSortedNode
m_nNodeCount = 0
Set m_objNode = Nothing
End Sub

Friend Sub AddEdge(ByVal objSrc As IASTNode, ByVal objDest As IASTNode)
On Error Resume Next
Dim s As String
Dim idxs As Long, idxe As Long
Dim tmp As Long
'///
Err.Clear
s = CStr(ObjPtr(objSrc))
idxs = m_objNode.Item(s)
If Err.Number Then
 m_nNodeCount = m_nNodeCount + 1
 ReDim Preserve m_tNodes(1 To m_nNodeCount)
 Set m_tNodes(m_nNodeCount).obj = objSrc
 '///
 idxs = m_nNodeCount
 m_objNode.Add idxs, s
End If
'///
Err.Clear
s = CStr(ObjPtr(objDest))
idxe = m_objNode.Item(s)
If Err.Number Then
 m_nNodeCount = m_nNodeCount + 1
 ReDim Preserve m_tNodes(1 To m_nNodeCount)
 Set m_tNodes(m_nNodeCount).obj = objDest
 '///
 idxe = m_nNodeCount
 m_objNode.Add idxe, s
End If
'///
tmp = m_tNodes(idxs).nDegreeOut + 1
m_tNodes(idxs).nDegreeOut = tmp
ReDim Preserve m_tNodes(idxs).nOut(1 To tmp)
m_tNodes(idxs).nOut(tmp) = idxe
'///
m_tNodes(idxe).nDegreeIn = m_tNodes(idxe).nDegreeIn + 1
End Sub

Friend Property Get NodeCount() As Long
NodeCount = m_nNodeCount
End Property

Friend Function SortedNode(ByVal nIndex As Long) As IASTNode
Set SortedNode = m_tNodes(m_nSortedNode(nIndex)).obj
End Function
