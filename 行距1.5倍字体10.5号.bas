Attribute VB_Name = "ģ��1"
Option Explicit
Sub NewBar()

    Dim bar As CommandBar
    
    Set bar = CommandBars.Add(Name:="���Թ�����")
    
    bar.Visible = True
    
    Dim but As CommandBarButton
    
    Set but = bar.Controls.Add(Type:=msoControlButton)
        
    With but
        
        .BeginGroup = True
            
        .TooltipText = "���԰�ť1"
            
        .Caption = "�ֺ�10.5�о�1.5��"
            
        .Style = msoButtonCaption
        
        .OnAction = "BatchSpacing"
        
    End With

End Sub
Sub BatchSpacing()  '���������ı��о༰�ֺ�

    Dim oSlides As Slides
    
    Dim oSlide As Slide
    
    Dim oPre As Presentation
    
    Dim oShape As Shape
    
    Dim oTextFrame As TextFrame
    
    Set oPre = ActivePresentation
    
    Set oSlides = oPre.Slides
    
    Dim i As Integer, j As Integer
    
    For i = 1 To oSlides.Count
    
        Set oSlide = oSlides.Item(i)
    
        For j = 1 To oSlide.Shapes.Count
        
            Set oShape = oSlide.Shapes(j)
            
            Set oTextFrame = oShape.TextFrame
            
            If oTextFrame.HasText Then
            
                '�����о�Ϊ1.5��
                oTextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.5
                
                oTextFrame.TextRange.Font.Size = 10.5
            
            End If
            
        Next
    
    Next

End Sub

Sub DelBar()
    DelToolBar "���Թ�����"
End Sub
Sub DelToolBar(ByVal barName As String)
    CommandBars(barName).Delete
End Sub
