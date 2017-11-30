(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // A função inicializar deverá ser executada cada vez que uma nova página for carregada.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // inicialize o mecanismo de notificação do FabricUI e oculte-o
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // Se não estiver usando o Excel 2016, use a lógica de fallback.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("Este exemplo exibirá o valor das células que você selecionou na planilha.");
                $('#button-text').text("Exibir!");
                $('#button-desc').text("Exibir a seleção");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("Este exemplo realça o valor mais alto das células que você selecionou na planilha.");
            $('#button-text').text("Realçar!");
            $('#button-desc').text("Realça o maior número.");
                
            loadSampleData();

            // Adicione um manipulador de eventos de clique ao botão de realce.
            $('#highlight-button').click(hightlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Executar uma operação em lote com base no modelo de objeto do Excel
        Excel.run(function (ctx) {
            // Crie um objeto de proxy para a variável de planilha ativa
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Colocar um comando na fila para gravar os dados de exemplo na planilha
            sheet.getRange("B3:D5").values = values;

            // Executar os comandos na fila e retornar uma promessa para indicar a conclusão da tarefa
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function hightlightHighestValue() {
        // Executar uma operação em lote com base no modelo de objeto do Excel
        Excel.run(function (ctx) {
            // Criar um objeto proxy para o intervalo selecionado e carregar suas propriedades
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Executar o comando na fila e retornar uma promessa para indicar a conclusão da tarefa
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Localizar a célula para realçar
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Realçar a célula
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('O texto selecionado é:', '"' + result.value + '"');
                } else {
                    showNotification('Erro', result.error.message);
                }
            });
    }

    // Função auxiliar para tratar erros
    function errorHandler(error) {
        // Sempre se certifique de capturar erros acumulados que surgirem da execução do Excel.run
        showNotification("Erro", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Função auxiliar para exibir notificações
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
