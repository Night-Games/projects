async function insertYouTubeVideo() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            const shape = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
            shape.left = 50; // posição X inicial (em pixels)
            shape.top = 50;  // posição Y inicial (em pixels)
            shape.width = 560;
            shape.height = 315;

            // Código para o iframe
            const iframeHTML = `
                <iframe width="560" height="315"
                    src="https://www.youtube.com/embed/Re4FlYXe6Cg?autoplay=1"
                    frameborder="0" allow="autoplay; encrypted-media" allowfullscreen>
                </iframe>
            `;

            shape.textFrame.textRange.text = ""; // Limpa o texto
            shape.fill.setTransparency(1); // Deixa a forma transparente
            shape.lineFormat.visible = false; // Remove a borda

            // Usa o "alt text" para tentar simular (não é ideal, mas ajuda para hover/identificação)
            shape.altTextDescription = "YouTube Video";

            // Inserir o vídeo de outra forma é bloqueado diretamente (por segurança).
            // Se quiser tentar evolução depois, podemos melhorar ainda mais.

            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}