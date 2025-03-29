const fs = require("fs-extra");
const path = require("path");
const csv = require("csv-parser");
const xlsx = require("xlsx");
const puppeteer = require("puppeteer");
const { exec } = require("child_process");
const { MailerSend } = require("mailersend");

const mailersend = new MailerSend({ apiKey: "TU_CLAVE_MAILERSEND" });
const templateId = "TEMPLATE_CREADO_MAILERSEND";
const inputFile = "Archivo.xlsx"; // o donaciones.csv
const htmlTemplate = "Plantilla.html"; // En caso de ser un formato HTML
const logFile = "log-envios.xlsx"; // log de envío
const pdfOutputDir = "pdfs_enviados"; // carpeta donde genera los archivos

const HEADERS_REQUERIDOS = ["NOMBRE", "NIF", "IMPORTE", "EMAIL"];
const SOLO_REINTENTAR_FALLIDOS = false;

fs.ensureDirSync(pdfOutputDir);

function limpiarNombreArchivo(nombre) {
  return nombre.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\w\d]/g, "_");
}

function leerLogExistente() {
  if (!fs.existsSync(logFile)) return [];
  const workbook = xlsx.readFile(logFile);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
}

function guardarLog(datos) {
  const ws = xlsx.utils.json_to_sheet(datos);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Enviados");
  xlsx.writeFile(wb, logFile);
}

async function leerArchivo(pathArchivo) {
  const ext = path.extname(pathArchivo).toLowerCase();
  if (ext === ".csv") return leerCSV(pathArchivo);
  if (ext === ".xlsx" || ext === ".xls") return leerExcel(pathArchivo);
  throw new Error("Formato de archivo no soportado: " + ext);
}

function leerCSV(pathArchivo) {
  return new Promise((resolve, reject) => {
    const resultados = [];
    fs.createReadStream(pathArchivo)
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

function leerExcel(pathArchivo) {
  const wb = xlsx.readFile(pathArchivo, { cellText: false, cellNF: false });
  const hoja = wb.Sheets[wb.SheetNames[0]];
  // raw: true => devuelve valores numéricos genuinos en lugar de strings parseadas
  const datos = xlsx.utils.sheet_to_json(hoja, { defval: "", raw: true });
  const faltantes = HEADERS_REQUERIDOS.filter(h => !Object.keys(datos[0] || {}).includes(h));
  if (faltantes.length) throw new Error("Faltan columnas: " + faltantes.join(", "));

  // Aquí podrías formatear cada importe en el formato ES antes de devolver el array completo
  for (const row of datos) {
    if (typeof row.IMPORTE === "number") {
      // Con Intl.NumberFormat en español => "1.220,45"
      row.IMPORTE = new Intl.NumberFormat("es-ES", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      }).format(row.IMPORTE);
    }
  }

  return datos;
}


async function generarPDF(html, outputPath) {
  const browser = await puppeteer.launch({ headless: true, args: ["--no-sandbox", "--disable-setuid-sandbox"] });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle0" });
  await page.emulateMediaType("screen");
  await page.pdf({ path: outputPath, format: "A4" });
  console.log("✅ PDF generado:", outputPath);
  await browser.close();
}

async function crearZipConPassword(pdfPath, zipPath, password) {
  const comando = `"C:\\Program Files\\7-Zip\\7z.exe" a -tzip -p${password} -mem=AES256 "${zipPath}" "${pdfPath}"`;
  //const comando = `7z a -tzip -p${password} -mem=AES256 "${zipPath}" "${pdfPath}"`; // para Linux
  return new Promise((resolve, reject) => {
    exec(comando, (error, stdout, stderr) => {
      if (error) {
        reject(new Error(`Error al crear ZIP: ${stderr || stdout}`));
      } else {
        resolve();
      }
    });
  });
}

async function enviarCorreos() {
  try {
    const logExistente = leerLogExistente();
    const enviadosOK = new Set(logExistente.filter(x => x.Estado === "✅ Enviado Correctamente").map(x => x.Email));
    const fallidos = new Set(logExistente.filter(x => x.Estado === "❌ Error").map(x => x.Email));
    const entradas = await leerArchivo(inputFile);
    const logActualizado = [...logExistente];
    const htmlBase = fs.readFileSync(htmlTemplate, "utf8");

    console.log("Entradas leídas:", entradas.length);

    for (const row of entradas) {
      const { EMAIL, NOMBRE, NIF, IMPORTE } = row;

      if (!EMAIL || !NOMBRE || !NIF || !IMPORTE || !EMAIL.includes("@")) {
        console.warn(`⛔ Datos incompletos o email inválido: ${JSON.stringify(row)}`);
        logActualizado.push({
          Fecha: new Date().toISOString(),
          Nombre: NOMBRE || "(sin nombre)",
          NIF: NIF || "(sin NIF)",
          Importe: IMPORTE || "(sin importe)",
          Email: EMAIL || "(sin email)",
          Estado: "❌ Error",
          Error: "Datos incompletos o email inválido"
        });
        continue;
      }
      

            if (!SOLO_REINTENTAR_FALLIDOS && enviadosOK.has(EMAIL)) {
        console.log(`⏭ Ya enviado correctamente a ${EMAIL}, saltando...`);
        continue;
      }

      if (SOLO_REINTENTAR_FALLIDOS && !fallidos.has(EMAIL)) continue;

     
      const htmlPersonalizado = htmlBase
      .replace(/{{\s*NOMBRE\s*}}/g, NOMBRE)
      .replace(/{{\s*NIF\s*}}/g, NIF)
      .replace(/{{\s*IMPORTE\s*}}/g, IMPORTE);

      const nombreLimpio = limpiarNombreArchivo(NOMBRE);
      const pdfPath = path.join(pdfOutputDir, `recibo-${nombreLimpio}.pdf`);
      const zipPath = path.join(pdfOutputDir, `recibo-${nombreLimpio}.zip`);

      try {
        await generarPDF(htmlPersonalizado, pdfPath);
        await crearZipConPassword(pdfPath, zipPath, NIF);

        const zipBase64 = fs.readFileSync(zipPath).toString("base64");

        const emailParams = {
          from: { email: "correo_envio", name: "titulo" },
          to: [{ email: EMAIL, name: NOMBRE }],
          subject: "Asunto",
          template_id: templateId,
          variables: [
            {
              email: EMAIL,
              substitutions: [
                { var: "NOMBRE", value: String(NOMBRE) },
                { var: "NIF", value: String(NIF) },
                { var: "IMPORTE", value: String(IMPORTE) }
            ]
            }
          ],
          attachments: [
            {
              filename: `recibo-${nombreLimpio}.zip`,
              content: zipBase64,
              type: "application/zip"
            }
          ]
        };

        try {
          console.log(`📤 Enviando a ${EMAIL}...`);
          console.log("📎 Adjuntos:", emailParams.attachments[0].filename);
          console.log("📧 Variables:", JSON.stringify(emailParams.variables, null, 2)); // Muestra el arreglo que retorna

          await mailersend.email.send(emailParams);
          console.log(`✅ Enviado a ${EMAIL}`);

          logActualizado.push({
            Fecha: new Date().toISOString(),
            Nombre: row.NOMBRE,
             NIF: row.NIF,
             Importe: row.IMPORTE,
             Email: row.EMAIL,
             Estado: "✅ Enviado Correctamente",
             Error: ""
          });
          // Si activa se eliminan PDF y ZIP
          //fs.unlinkSync(pdfPath);
          //fs.unlinkSync(zipPath);
        } catch (err) {
          console.error(`❌ Error con ${EMAIL}:`);

          if (err?.response) {
            console.error("🛑 Código de estado:", err.response.statusCode);
            console.error("📩 Respuesta completa:", JSON.stringify(err.response.body, null, 2));
          } else if (err?.message) {
            console.error("🛑 Mensaje de error:", err.message);
          } else {
            console.error("🛑 Error desconocido:", err);
          }

          logActualizado.push({
            Fecha: new Date().toISOString(),
            Nombre: NOMBRE,
            NIF,
            Importe: IMPORTE,
            Email: EMAIL,

            Estado: "❌ Error",
            Error: err?.response?.body ? JSON.stringify(err.response.body) : err?.message || "Error desconocido"
          });
        }
      } catch (err) {
        console.error("❌ Error en la generación del archivo para ${EMAIL}:", err.message);
        logActualizado.push({
          Fecha: new Date().toISOString(),
          Nombre: NOMBRE,
          NIF,
          Importe: IMPORTE,
          Email: EMAIL,
          Estado: "❌ Error Generando PDF/ZIP",
          Error: err.message
        });
      }
    }

    guardarLog(logActualizado);
    console.log("📊 Log actualizado en", logFile);
  } catch (error) {
    console.error("❌ Error general:", error.message);
  }
}

enviarCorreos();