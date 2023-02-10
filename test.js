const fs = require("fs");
const puppeteer = require("puppeteer");
const officegen = require("officegen");

async function captureScreen(link) {
  // Ouvrir le navigateur et créer une nouvelle page
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
 // Vérifier si une fenêtre de cookie est affichée et la fermer si nécessaire
 const cookieWindow = await page.$(".cookie-banner");
 if (cookieWindow) {
   await cookieWindow.click("button.cookie-banner__btn");
 }
  // Aller à la page désignée par le lien
  await page.goto(link);

  // Zoom arrière de 25%
  await page.evaluate(() => {
    document.body.style.zoom = "0.50";
  });

  // Cliquer sur la flèche bas
  await page.keyboard.press("ArrowDown");

  // Prendre une capture d'écran de la page entière avec une résolution de 1020 x 1020
  const screenshot = await page.screenshot({ clip: { x: 0, y: 0, width: 1200, height: 800 } });

  // Fermer le navigateur
  await browser.close();

  return screenshot;
}

async function run() {
  // Lire la liste des liens à partir d'un fichier
  const links = fs.readFileSync("links.txt").toString().split("\n");

  // Créer un nouveau fichier PowerPoint
  const pptx = officegen("pptx");

  // Boucle sur chaque lien dans la liste
  for (const link of links) {
    // Ignorer les lignes vides
    if (!link) continue;

    // Prendre une capture d'écran de la page associée au lien
    const screenshot = await captureScreen(link);

    // Ajouter une nouvelle diapositive à la fin de la présentation PowerPoint
    const slide = pptx.makeNewSlide();

    // Ajouter la capture d'écran à la diapositive à gauche
    slide.addImage(screenshot, { x: 0, y: 0, cx: "80%", cy: "80%" });

    // Ajouter le nom du site en gras à gauche, la date et le lien comprimé
    const siteName = link.replace("https://www.", "").split(".")[0];
const date = new Date().toLocaleDateString();
const compressedLink = link;
const title = slide.addText(siteName);
title.options.align = "right";
title.options.bold = true;
slide.addText(date, { x: "50%", y: 0, cx: "50%", cy: "20%", align: "right" });
slide.addText(compressedLink, { x: "50%", y: "20%", cx: "50%", cy: "20%", url: link, align: "right" });
slide.addText(siteName, { x: "50%", y: 0, cx: "50%", cy: "20%", align: "right" });

}
  // Enregistrer le fichier PowerPoint
  pptx.generate(fs.createWriteStream("presentation.pptx"));
}

run();
