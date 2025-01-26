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
    // Appel de la fonction pour écouter les événements clavier
    listenForKeyPress();

    // Fonction pour écouter l'appui sur une touche spécifique
    function listenForKeyPress() {
      document.addEventListener('keydown', function(event) {
    // Vérifie si la touche pressée est une lettre
        if (event.key.length === 1 && /[a-zA-Z]/.test(event.key)) {
      // Appel à la fonction pour changer la couleur de la cellule active
          changeCellColor();
        }
      });
    }

    // Fonction pour changer la couleur de la cellule active
    async function changeCellColor() {
      try {
        await Excel.run(async (context) => {
          // Obtient la cellule active
          const activeCell = context.workbook.getActiveCell();
          // Modifie la couleur de fond de la cellule active (par exemple, rouge)
          activeCell.format.fill.color = "red"; 
          await context.sync();
        });
      } catch (error) {
        console.error("Erreur lors de la modification de la couleur : ", error);
      }
    }



  } else {
    console.error("This add-in is not running in Excel.");
  }
});
