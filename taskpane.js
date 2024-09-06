Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        console.log("Excel add-in loaded");
        // Aqui você pode adicionar mais código para interagir com o Excel, se necessário
    } else {
        console.log("This add-in is not running in Excel");
    }
});
