Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office.js is ready.");

    // Appel de la fonction pour écouter les événements clavier
    listenForKeyPress();
    
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

// Fonction pour écouter les appuis clavier
    function listenForKeyPress() {
      document.addEventListener("keydown", (event) => {
    console.log(`Key pressed: ${event.key}`); // Log l'appui de touche
    if (event.key === "a") { // Si la touche "a" est pressée
      changeCellColor(); // Appelle la fonction pour changer la couleur de la cellule
    }
  });
    }

// Fonction pour changer la couleur de la cellule active
    async function changeCellColor() {
      try {
        await Excel.run(async (context) => {
          console.log("Changing the color of the active cell...");
          const activeCell = context.workbook.getActiveCell();
      activeCell.format.fill.color = "yellow"; // Couleur de fond modifiée
      await context.sync();
      console.log("Cell color changed to yellow.");
    });
      } catch (error) {
        console.error("Error changing cell color: ", error);
      }



    } else {
      console.error("This add-in is not running in Excel.");
    }
  });
