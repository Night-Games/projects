Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        Office.context.ui.addHandlerAsync(Office.EventType.DialogMessageReceived, function (event) {
            if (event.message === 'addVideo') {
                addContentAppToSheet();
            }
        });
    }
});

async function addContentAppToSheet() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange("A1"); // Você pode definir onde deseja adicionar o conteúdo
            range.values = [["<iframe width='560' height='315' src='https://www.youtube.com/embed/CjP3VlbIfDA' frameborder='0' allowfullscreen></iframe>"]];
            await context.sync();
        });
    } catch (error) {
        console.error('Erro ao adicionar conteúdo: ', error);
    }
}
