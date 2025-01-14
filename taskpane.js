document.getElementById("createSheet").addEventListener("click", () => {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add("New Sheet");
    sheet.activate();
    await context.sync();
    console.log("New sheet created!");
  }).catch((error) => {
    console.error(error);
  });
});
