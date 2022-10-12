REM  *****  BASIC  *****
Option Explicit

Const MY_LIBRARY = "Standard", MY_DIALOG = "Dialog1", MY_BUTTON = "Button1"

Const MY_LABEL = "Inserir Imagem"

Sub Main
    Dim libr As Object ' com.sun.star.script.XLibraryContainer
    Dim dlg As Object

    Dim ui As Object  ' stardiv.Toolkit.UnoDialogControl

    Dim ctl As Object ' stardiv.Toolkit.UnoButtonControl 

    Dim rc As Object : rc = com.sun.star.ui.dialogs.ExecutableDialogResults

    BasicLibraries.LoadLibrary(MY_LIBRARY)
    libr = DialogLibraries.GetByName(MY_LIBRARY)
   
    dlg = libr.GetByName(MY_DIALOG)

    
    ui = CreateUnoDialog(dlg)

    
    ui.Title = "Título Qualquer EXEMPLO 2022"

 
    ctl = ui.GetControl(MY_BUTTON)
    ctl.Model.Label = MY_LABEL

    
    Select Case ui.Execute
        Case rc.OK 
        	setPerDialog
        Case rc.CANCEL
        	MsgBox "O usuário cancelou o diálogo.", 0 , "Basic"
    End Select
    ui.endExecute()
    

End Sub

Sub insertPerDialog
	
	On Error GoTo ErrorHandler 'Tratamento de exceções
	
	dim oDP as object
	dim oObj as object
	dim imgDescription as String 'String que recebe descrição da imagem
	dim i as Integer 'Contador imagens

	i = 0 'posição inicial do vetor de imagens
	
	oDP = ThisComponent.DrawPage
	oObj = oDP.getByIndex(i) 'oObj recebe o indice da primeira imagem, posição 0
	
	'print  (isEmpty(oObj))

	do while (not (isEmpty(oObj))) 'faça enquanto vazio é falso == posição cheia
	
			i = i + 1 'incrementa index 
			
			oObj = oDP.getByIndex(i) 'proxima posição
				
	loop
	
		ErrorHandler: 'se der erro vem para cá
	'	MsgBox "Valor vazio " & i 'Mostra qual o proximo valor vazio
	
		InserirImagem 'chama a sub 
		
		imgDescription = InputBox("Descrição da imagem") 'Espera do teclado a descrição da imagem'	
		oObj = oDP.getByIndex(i) 'Index = pos da imagem no Array
		oObj.Description = imgDescription 'altera a descrição da imagem
		
		MsgBox "Operação finalizada"
		End Sub
	

Sub InserirImagem
rem ----------------------------------------------------------------------
rem define variables
	On Error GoTo ErrorHandl

	Dim document   as object
	Dim dispatcher as object
	Dim imageURL 

rem ----------------------------------------------------------------------
rem get access to the document

	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
	
	dispatcher.executeDispatch(document, ".uno:InsertGraphic", "", 0, Array()) 'Insere Imagem
	
	'	imgDescription = InputBox("Descrição da imagem") 'Espera do teclado a descrição da imagem
	'	oObj = oDP.getByIndex(i) 'Index = pos da imagem no Array
		'oObj.Description = imgDescription 'altera a descrição da imagem

'	dispatcher.executeDispatch(document, ".uno:SelectObject", "", 0, Array())
	'dispatcher.executeDispatch(document, ".uno:GraphicDialog", "", 0, Array())
	'dispatcher.executeDispatch(document, ".uno:InsertAnnotation", "", 0, Array()) 'Insere anotação
ErrorHandl:
End Sub

Sub createTxt

	Dim Doc As Object
	Dim Enum As Object
	Dim TextElement As Object
	Dim iCount
	Dim Path as String
	
	iCount = Freefile
	Path = InputBox("Especifique o caminho do Arquivo(Exemplo: C:\Users\User\Desktop\NOME_DO_ARQUIVO.txt)")
	open Path for OutPut as iCount
	'open "C:\users\War Machine\desktop\data.txt" for OutPut as iCount
	
	 
	Doc = ThisComponent
	Enum = Doc.Text.createEnumeration
	 
	While Enum.hasMoreElements
	
	'	MsgBox TextElement.String
	
	  TextElement = Enum.nextElement
	 
	  If TextElement.supportsService("com.sun.star.text.Paragraph") Then
	    'TextElement.String = Replace(TextElement.String, "you", "U") 
	    'TextElement.String = Replace(TextElement.String, "too", "2")
	    'TextElement.String = Replace(TextElement.String, "for", "4")
	    
	    Write #iCount, TextElement.String
	     
	    ' MsgBox TextElement.String
	     
	  End If
	 
	Wend
		MsgBox "Arquivo txt criado"

End Sub
