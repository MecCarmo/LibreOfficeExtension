Sub setPerDialog
	
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
		MsgBox "Valor vazio " & i 'Mostra qual o proximo valor vazio
	
		InserirImagem 'chama a sub 
		
		imgDescription = InputBox("Descrição da imagem") 'Espera do teclado a descrição da imagem'	
		oObj = oDP.getByIndex(i) 'Index = pos da imagem no Array
		oObj.Description = imgDescription 'altera a descrição da imagem
		
		MsgBox "Operação finalizada"
		End Sub
