// Requiere: npm install mailersend csv-parser xlsx puppeteer
const fs = require("fs");
const csv = require("csv-parser");
const xlsx = require("xlsx");
const puppeteer = require("puppeteer");
const { MailerSend } = require("mailersend");

const mailersend = new MailerSend({ apiKey: "TU_CLAVE_MAILERSEND" });
const templateId = "TEMPLATE_CREADO_MAILERSEND";
const inputFile = "Archivo.xlsx"; // o donaciones.csv
const htmlTemplate = "Plantilla.html"; // En caso de ser un formato HTML
const logFile = "log-envios.xlsx"; // log de envío
const pdfOutputDir = "pdfs_enviados"; // carpeta donde genera los archivos

function leerLogExistente() {
  if (!fs.existsSync(logFile)) return [];
  const workbook = xlsx.readFile(logFile);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
}

function guardarLog(datos) {
  const ws = xlsx.utils.json_to_sheet(datos);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Enviados Correctamente");
  xlsx.writeFile(wb, logFile);
}

function leerCSV(path) {
  return new Promise((resolve, reject) => {
    const resultados = [];
    fs.createReadStream(path)
      .pipe(csv())
      .on("headers", headers => {
        const faltantes = HEADERS_REQUERIDOS.filter(h => !headers.includes(h));
        if (faltantes.length) reject(new Error("Faltan columnas: " + faltantes.join(", ")));
      })
      .on("data", data => resultados.push(data))
      .on("end", () => resolve(resultados))
      .on("error", reject);
  });
}

async function generarPDF(html, outputPath) {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle0" });
  await page.pdf({ path: outputPath, format: "A4" });
  await browser.close();
}

async function enviarCorreos() {
  try {
    const logExistente = leerLogExistente();
    const enviados = new Set(logExistente.filter(x => x.Estado === "✅ Enviado Correctamente").map(x => x.Email));
    const entradas = await leerCSV(inputFile);
    const logActualizado = [...logExistente];
    const htmlBase = fs.readFileSync(htmlTemplate, "utf8");

    for (const row of entradas) {
      const email = row.EMAIL;
      if (enviados.has(email)) {
        console.log(`⏭ Ya fue enviado correctamente a ${email}, saltando...`);
        continue;
      }

  // Esto permite que haya espacios dentro del {{ }} y lo hace insensible a ese detalle
        const htmlPersonalizado = htmlBase
        .replace(/{{\s*NOMBRE\s*}}/g, row.NOMBRE)
        .replace(/{{\s*NIF\s*}}/g, row.NIF)
        .replace(/{{\s*IMPORTE\s*}}/g, row.IMPORTE);

      const pdfPath = `recibo-${row.NIF}.pdf`;
      await generarPDF(htmlPersonalizado, pdfPath);
      const pdfBase64 = fs.readFileSync(pdfPath).toString("base64");

      const emailParams = {
        from: { email: "correo_envio", name: "titulo" },
        to: [{ email: EMAIL, name: NOMBRE }],
        subject: "Asunto",
        template_id: templateId,
        variables: [
          {
            email: email,
            substitutions: [
              { var: "NOMBRE", value: row.NOMBRE },
              { var: "NIF", value: row.NIF },
              { var: "IMPORTE", value: row.IMPORTE }
            ]
          }
        ],
        attachments: [
          {
            filename: `recibo-${row.NIF}.pdf`,
            content: pdfBase64,
            type: "application/pdf"
          }
        ]
      };

      try {
        await mailersend.email.send(emailParams);
        console.log(`✅ Enviado Correctamente a ${email}`);
        logActualizado.push({
          Fecha: new Date().toISOString(),
          Nombre: row.NOMBRE,
          NIF: row.NIF,
          Importe: row.IMPORTE,
          Email: email,
          Estado: "✅ Enviado Correctamente",
          Error: ""
        });
      } catch (err) {
        console.error(`❌ Error con ${email}:`, err.message);
        logActualizado.push({
          Fecha: new Date().toISOString(),
          Nombre: row.NOMBRE,
          NIF: row.NIF,
          Importe: row.IMPORTE,
          Email: email,
          Estado: "❌ Error",
          Error: err.message
        });
      }
    }

    guardarLog(logActualizado);
    console.log("📊 Log actualizado en log-envios.xlsx");
  } catch (error) {
    console.error("❌ Error general:", error.message);
  }
}

enviarCorreos();
