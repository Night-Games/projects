Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    // Office is ready
    // Event handler for button click
    $("#video").on("click", run);
  }
});

async function run() {
  // Display a dialog to show the HTML content
  Office.context.ui.displayDialogAsync(
    "https://ecogs2022.github.io/Testes/taskpane.html", // Correct URL
    { height: 500, width: 800 }, // Adjust size as needed
    async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        // Optionally handle dialog events if needed
        dialog.addHandlerAsync(Office.EventType.DialogMessageReceived, (event) => {
          if (event.message === "close") {
            dialog.closeAsync();
          }
        });
      } else {
        console.error("Error displaying dialog: ", result.error.message);
      }
    }
  );
}
