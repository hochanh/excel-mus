Attribute VB_Name = "Module2"
Sub Lay_mau()
'
' Macro lay mau kich co n vao DS mau
'
' Keyboard Shortcut: Ctrl+s
'


' Xac nhan thuc hien chon danh sach mau
Dim Xacnhan
Xacnhan = MsgBox("Qua trinh nay se xoa danh sach mau da lap (neu co)." & vbNewLine & "Van tiep tuc?", vbExclamation + vbYesNo, "CHÚ Ý!")
If Xacnhan = 7 Then Exit Sub


Sheets("DS mau").Select

' Thiet dat gia tri cho cac bien
  
    Dim rngCell1 As Range, numCell As Range

    Set rngCell1 = Range("AA1")
    Set numCell = Range("AA2")

    Columns("A:AC").AutoFit
    Columns("A:AC").Select
    Selection.ClearContents
    
    Counter = 0

'Xac dinh khoang cach mau
    
    Sheets("Tao mau").Select
    PopSize = Range("F5")
    SmpSize = Range("F22")
    
    If SmpSize = 0 Then
    MsgBox "Dieu chinh cac thong so dau vao de giam co mau!", vbCritical + vbOKOnly, "CO MAU QUA LON!"
    Exit Sub
    End If
        
    Interval = PopSize / SmpSize
    Lower = 0
    
'Xac dinh diem bat dau ngau nhien

    Sheets("DS mau").Select
    Randomize
    Beg = Int((Interval - Lower + 1) * Rnd + Lower)
    rngCell1.Value = Beg
    Set rngCell1 = rngCell1.Offset(rowoffset:=1)
 
 
 'Tao mau don vi tien te
   
    Counter = 1
    Do Until Beg = PopSize Or Beg > PopSize Or Counter = SmpSize
        Beg = Beg + Interval
        rngCell1.Value = Beg
        Set rngCell1 = rngCell1.Offset(rowoffset:=1)
        Counter = Counter + 1
    Loop


'Don dep va tao DS mau
   
   Sheets("DS mau").Select
    Columns("A:L").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = True
        .MergeCells = False
    End With
    


'De xac dinh gia tri tien te
 
    Sheets("DS mau").Select
    Range("Z1").Value = 1
        'Set numCell = Range("AA2")
        CntCell = 1
        Do Until numCell.Value = ""
            CntCell = CntCell + 1
            numCell.Offset(columnoffset:=-1).Value = CntCell
            Set numCell = numCell.Offset(rowoffset:=1)
        Loop
    
    
'Sao chep gia tri tien te vao DS mau
         
            
    'Copies first 50 dollar items
        
        Sheets("DS mau").Select
        Range("Z1:AA50").Select
        Selection.Copy
        Range("A2").Select
        ActiveSheet.Paste
        Range("A1").Value = "#"
        Range("B1").Value = "Gia tri bang tien"
        Range("C1").Value = "Khoan muc tuong ung"
        Range("D1").Value = "Co sai sot?"
    
    'Copies sample items 51 - 100 if needed
    
        If CntCell > 50 Then
            Range("Z51:AA100").Select
            Selection.Copy
            Range("F2").Select
            ActiveSheet.Paste
            Range("F1").Value = "#"
            Range("G1").Value = "Gia tri bang tien"
            Range("H1").Value = "Khoan muc tuong ung"
            Range("I1").Value = "Co sai sot?"
        End If
    
    'Copies sample items 101 - 150 if needed
        
        If CntCell > 100 Then
            Range("Z101:AA150").Select
            Selection.Copy
            Range("A53").Select
            ActiveSheet.Paste
            Range("A52").Value = "#"
            Range("B52").Value = "Gia tri bang tien"
            Range("C52").Value = "Khoan muc tuong ung"
            Range("D52").Value = "Co sai sot?"
        End If
    
    'Copies sample items 151 - 200 if needed
        
        If CntCell > 150 Then
            Range("Z151:AA200").Select
            Selection.Copy
            Range("F53").Select
            ActiveSheet.Paste
            Range("F52").Value = "#"
            Range("G52").Value = "Gia tri bang tien"
            Range("H52").Value = "Khoan muc tuong ung"
            Range("I52").Value = "Co sai sot?"
        End If
    
    
    'Copies sample items 201 - 250 if needed
        
        If CntCell > 200 Then
            Range("Z201:AA250").Select
            Selection.Copy
            Range("A104").Select
            ActiveSheet.Paste
            Range("A103").Value = "#"
            Range("B103").Value = "Gia tri bang tien"
            Range("C103").Value = "Khoan muc tuong ung"
            Range("D103").Value = "Co sai sot?"
        End If
    
        
    'Copies sample items 251 - 300 if needed
        
        If CntCell > 250 Then
            Range("Z251:AA300").Select
            Selection.Copy
            Range("F104").Select
            ActiveSheet.Paste
            Range("F103").Value = "#"
            Range("G103").Value = "Gia tri bang tien"
            Range("H103").Value = "Khoan muc tuong ung"
            Range("I103").Value = "Co sai sot?"
        End If


    'Copies sample items 301 - 350 if needed
    
        If CntCell > 300 Then
            Range("Z301:AA350").Select
            Selection.Copy
            Range("A155").Select
            ActiveSheet.Paste
            Range("A154").Value = "#"
            Range("B154").Value = "Gia tri bang tien"
            Range("C154").Value = "Khoan muc tuong ung"
            Range("D154").Value = "Co sai sot?"
        End If
    
        
    'Copies sample items 351 - 400 if needed
        
        If CntCell > 350 Then
            Range("Z351:AA400").Select
            Selection.Copy
            Range("F155").Select
            ActiveSheet.Paste
            Range("F154").Value = "#"
            Range("G154").Value = "Gia tri bang tien"
            Range("H154").Value = "Khoan muc tuong ung"
            Range("I154").Value = "Co sai sot?"
        End If
   
    'Copies sample items 401 - 450 if needed
    
        If CntCell > 400 Then
            Range("Z401:AA450").Select
            Selection.Copy
            Range("A206").Select
            ActiveSheet.Paste
            Range("A205").Value = "#"
            Range("B205").Value = "Gia tri bang tien"
            Range("C205").Value = "Khoan muc tuong ung"
            Range("D205").Value = "Co sai sot?"
        End If
    
        
    'Copies sample items 451 - 500 if needed
        
        If CntCell > 450 Then
            Range("Z451:AA500").Select
            Selection.Copy
            Range("F206").Select
            ActiveSheet.Paste
            Range("F205").Value = "#"
            Range("G205").Value = "Gia tri bang tien"
            Range("H205").Value = "Khoan muc tuong ung"
            Range("I205").Value = "Co sai sot?"
        End If
        
'To format print page

    Range("B:B,G:G").Select
    Range("G1").Activate
    Selection.Style = "Currency"
    Selection.NumberFormat = "_(#,##0_);_((#,##0);_(""-""??_);_(@_)"
    Range("C:C,H:H").Select
    Range("H1").Activate
    Selection.ColumnWidth = 11
    Range("D:D,I:I").Select
    Range("I1").Activate
    Selection.ColumnWidth = 14
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .MergeCells = False
    End With
    Columns("F:F").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .MergeCells = False
    End With
  

'To clear and set print sheet to focal point
 
    MsgBox "Qua trinh tao danh sach mau da hoan tat!" & vbNewLine & "Su dung danh sach mau nay de tien hanh kiem toan chi tiet.", vbInformation + vbOKOnly, "Hoàn thành!"
    Columns("Y:AC").Select
    Selection.ClearContents
    Range("A1").Select
    

End Sub
