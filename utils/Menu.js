//NOTE Tudo relacionado com o menu para execução no spreadsheet

function abrirMenu() {
	SpreadsheetApp.getUi()
		.createMenu('Gerar Apuração')
		.addItem('Iniciar', 'main')
		.addSeparator()
		.addItem("Apagar todas as tabelas - (Menos Filtros de titulos e Dashboard)", "removeAllTable")
		.addToUi();
}