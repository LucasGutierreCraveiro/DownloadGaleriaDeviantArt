; Desenvolvido por Lucas Gutierre Craveiro
; 04/08/2018 
; O objetivo desse script é realizar o download de todas as imagens
; de uma galeria no DeviantArt.


#SingleInstance force

;Entrada de Dados
;MsgBox "Atenção! Caso a galeria possua imagens +18 é necessário estar logado no DeviantArt pelo !!!INTERNET EXPLORER!!!!. Caso o contrário ocorrerá um erro no script"
InputBox galeriaDeviantArt, "Download de Galeria", "Insira o link da galeria: "
InputBox caminhoDownload, "Destino da Imagem", "Insira o local aonde deseja salvar a imagem: "
MsgBox "O script realizará o download de todas as imagens da galeria inserida, para encerrar o script a qualquer momento pressione Ctrl + X "


pwb := ComObjCreate("InternetExplorer.Application") ; Cria o Internet Explorer
pwb.visible:=true  ; Deixe como "True" caso queira visualizar o Internet Explorer realizando as funções dele 

pwb.Navigate(galeriaDeviantArt) ;Vai pra página da galeria 


while pwb.busy or pwb.ReadyState != 4 ; Espera carregar
	Sleep, 100

while (pwb.document.getElementsByClassName("torpedo-thumb-link")[0].outerHTML ="")
	Sleep, 200

tamanhoGaleria := pwb.document.getElementsByClassName("torpedo-thumb-link").Length

pwb.document.all.tags("div")[0].fireEvent("scroll")  ; Sometimes needed in conjunction with changing value



i = 0

While i < tamanhoGaleria {

	while pwb.busy or pwb.ReadyState != 4 ; Espera carregar
	Sleep, 100

	pwb.document.getElementsByClassName("torpedo-thumb-link")[i].click()

	while pwb.busy or pwb.ReadyState != 4 ; Espera carregar
		Sleep, 100

	linkParaDownload := pwb.document.getElementsByClassName("dev-content-normal")[0].src

	;MsgBox % linkParaDownload ;;DEBUG

	pwb.Navigate(linkParaDownload) ; Navega até a imagem

	/* 
		Não é necessário navegar até o site da imagem já o script conseguiu capturar acima o link do download da imagem
		Porém eu acho melhor ir até o link para evitar que o IE dê crash
	*/


	; Faz o download da imagem, salva no local específicado pelo usuário com o nome de i (0, 1, 2 ,3 etc.)
	UrlDownloadToFile, % linkParaDownload, % caminhoDownload . i . ".jpg"


	;Volta pra galeria, se preparando pra ir para a próxima imagem
	pwb.navigate(galeriaDeviantArt)

	pwb.document.all.tags("div")[0].fireEvent("scroll")

	while pwb.busy or pwb.ReadyState != 4 ; CARREGUE! GALERIA!
	Sleep, 100

	while (pwb.document.getElementsByClassName("torpedo-thumb-link")[i].outerHTML ="")
	Sleep, 200

	i++

}

; Apenas para fins estéticos da mensagem final
i := i + 1


; Script Concluído com Sucesso
MsgBox "Realizado o Download de " . %i% . "imagens da galeria: " .  %galeriaDeviantArt%



; Em caso de emergência
^x::
	MsgBox Script Interrompido pelo Operador
	ExitApp
	return