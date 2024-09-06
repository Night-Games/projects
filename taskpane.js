Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        // Verifique se o Office está pronto
        $("#addVideoButton").on("click", function() {
            addContentAppToSheet();
        });
    }
});

async function addContentAppToSheet() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1"); // Ajuste conforme necessário
            range.values = [["<iframe width='560' height='315' src='https://www.youtube.com/embed/CjP3VlbIfDA' frameborder='0' allowfullscreen></iframe>"]];
            await context.sync();
        });
    } catch (error) {
        console.error('Erro ao adicionar conteúdo: ', error);
    }
}
