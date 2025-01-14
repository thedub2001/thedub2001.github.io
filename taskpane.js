Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("combine-formulas-button").onclick = combineFormulas;
    }
});

async function combineFormulas() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getSelectedRange();
        range.load("address, formulasR1C1, hasFormula");

        await context.sync();

        if (!range.hasFormula) {
            console.log("La cellule sélectionnée n'a pas de formule.");
            return;
        }

        const combinedFormula = await getCombinedFormula(context, range.formulasR1C1[0][0]);
        range.formulasR1C1 = [[combinedFormula]];
        await context.sync();
    });
}

async function getCombinedFormula(context, formula) {
    const formulaRegex = /\b[A-Z]{1,3}[1-9][0-9]*\b/g;
    let combinedFormula = formula;
    let match;

    // Trouver toutes les références de cellules dans la formule
    while ((match = formulaRegex.exec(formula)) !== null) {
        let ref = match[0];
        let refRange = context.workbook.worksheets.getActiveWorksheet().getRange(ref);
        refRange.load("formulasR1C1, hasFormula");

        await context.sync();

        if (refRange.hasFormula) {
            let refFormula = await getCombinedFormula(context, refRange.formulasR1C1[0][0]);

            // Supprimer le signe "=" en trop dans la formule de référence
            refFormula = refFormula.replace("=", "");

            // Remplacer les références par les formules correspondantes
            combinedFormula = combinedFormula.replace(ref, `(${refFormula})`);
        }
    }

    // Ajouter le signe "=" devant la formule combinée
    combinedFormula = "=" + combinedFormula;

    return combinedFormula;
}

