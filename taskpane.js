Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office.js is ready.");



      const button = document.getElementById("createSheet");
      if (button) {
        console.log("Button found. Adding event listener.");
        button.addEventListener("click", () => {
          console.log("Button clicked. Running Excel code...");
          Excel.run(async (context) => {
            console.log("Creating a new sheet...");
            const sheet = context.workbook.worksheets.add("New Sheet");
            sheet.activate();
            await context.sync();
            console.log("New sheet created!");
          }).catch((error) => {
            console.error("Error in Excel.run:", error);
          });
        });
      } else {
        console.error("Button with ID 'createSheet' not found.");
      }
    
  } else {
    console.error("This add-in is not running in Excel.");
  }
});
