Sub fixedWidthImport() 
' 
' fixedWidthImport Macro 
' Imports a fixed width file using the specificed parameters 
' 
' Keyboard Shortcut: Ctrl+Shift+P 
' 
Set myFile = Application.FileDialog(msoFileDialogOpen) 
With myFile 
.Title = "Choose a fixed width text file to import." 
.AllowMultiSelect = False 
If .Show <> -1 Then 
Exit Sub 
End If 
FileSelected = .SelectedItems(1) 
End With 
 
Workbooks.OpenText fileName:= _
FileSelected, Origin:= _ 
xlMSDOS, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 1), _ 
Array(1, 1), Array(21, 1), Array(24, 1), Array(33, 1), Array(53, 1), Array(62, 1), _ 
Array(92, 1), Array(112, 1), Array(122, 1), Array(125, 1), Array(126, 1), _ 
Array(130, 1), Array(135, 1), Array(135, 1), Array(135, 1), Array(135, 1), _ 
Array(140, 1), Array(145, 1), Array(147, 1), Array(149, 1), Array(151, 1), _ 
Array(157, 1), Array(163, 1), Array(169, 1), Array(175, 1), Array(181, 1), _ 
Array(164, 1), Array(194, 1), Array(204, 1), Array(214, 1), Array(224, 1), _ 
Array(234, 1), Array(237, 1), Array(277, 1), Array(279, 1), Array(283, 1), _ 
Array(333, 1), Array(342, 1), Array(351, 1), Array(401, 1), Array(451, 1), _ 
Array(461, 1), Array(471, 1), Array(486, 1), Array(489, 1), Array(490, 1), _ 
Array(496, 1), Array(497, 1), Array(498, 1), Array(508, 1), Array(511, 1), _ 
Array(512, 1), Array(513, 1), Array(525, 1), Array(528, 1), Array(540, 1), _ 
Array(542, 1), Array(555, 1), Array(560, 1), Array(565, 1), Array(570, 1), _ 
Array(575, 1), Array(580, 1), Array(589, 1), Array(591, 1), Array(593, 1), _ 
Array(603, 1), Array(613, 1), Array(619, 1), Array(625, 1), Array(651, 1), _ 
Array(653, 1), Array(654, 1), Array(661, 1), Array(668, 1), Array(675, 1), _ 
Array(682, 1), Array(689, 1), Array(696, 1), Array(703, 1), Array(710, 1), _ 
Array(717, 1), Array(724, 1), Array(731, 1), Array(743, 1), Array(758, 1), _ 
Array(761, 1), Array(762, 1), Array(764, 1), Array(766, 1)), TrailingMinusNumbers:=True 
 
End Sub 
