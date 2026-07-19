import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import axios from "axios";
import { imageSize } from "image-size";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import ImageModule from "docxtemplater-image-module-free";
 
console.log("✅ ProposalPro DEMO server is running");
 
const app = express();
 
app.use(cors());
app.use(express.json({ limit: "20mb" }));
 
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
 
// ---------------------------------------------------------------------
// DEMO LOGIN
// This is a public demo build — do NOT reuse real client credentials
// here. Single shared demo account, easy to hand out to prospects.
// ---------------------------------------------------------------------
const USERS = {
  demo: "demo"
};
 
const TEMPLATE_PATH = path.join(__dirname, "templates", "template.docx");
 
function clean(value) {
  if (value === undefined || value === null || value === "undefined") return "";
  return String(value).trim();
}
 
app.post("/login", (req, res) => {
  const { username, password } = req.body;
 
  if (USERS[username] && USERS[username] === password) {
    return res.json({ success: true });
  }
 
  return res.json({ success: false });
});
 
function parsePrice(value) {
  return Number(String(value || "0").replace(/[^\d.,]/g, "").replace(",", ".")) || 0;
}
 
function normalizeDiscount(value) {
  const discount = Number(String(value).replace(",", ".")) || 0;
  return Math.min(100, Math.max(0, discount));
}
 
function getDiscountedPrice(price, discount) {
  return price - (price * discount / 100);
}
 
function normalizeTemplateXml(zip) {
  let xml = zip.file("word/document.xml").asText();
 
  if (xml.includes("{photo}") && !xml.includes("{%photo}")) {
    xml = xml.replaceAll("{photo}", "{%photo}");
  }
 
  zip.file("word/document.xml", xml);
 
  return zip;
}
 
function loadTemplateZip() {
  const content = fs.readFileSync(TEMPLATE_PATH, "binary");
  const zip = new PizZip(content);
  return normalizeTemplateXml(zip);
}
 
async function fetchImageBuffer(url) {
  const fixedUrl = url.replace("/upload/", "/upload/f_png/");
  const res = await axios.get(fixedUrl, {
    responseType: "arraybuffer",
    timeout: 20000,
  });
 
  return Buffer.from(res.data);
}
 
async function prefetchImages(products) {
  const cache = new Map();
 
  for (const p of products) {
    if (!p.photo) continue;
 
    try {
      const buffer = await fetchImageBuffer(p.photo);
      cache.set(p.photo, buffer);
    } catch (err) {
      console.warn("⚠️ Image error:", p.photo, err.message);
    }
  }
 
  return cache;
}
 
function buildImageModule(imageCache) {
  return new ImageModule({
    centered: false,
 
    getImage(tagValue) {
      return imageCache.get(tagValue) || Buffer.alloc(0);
    },
 
    getSize(img) {
      if (!img || !img.length) return [1, 1];
 
      try {
        const dim = imageSize(img);
        const max = 140;
 
        let w = dim.width || max;
        let h = dim.height || max;
 
        const scale = max / Math.max(w, h);
 
        if (scale < 1) {
          w = Math.round(w * scale);
          h = Math.round(h * scale);
        }
 
        return [w, h];
      } catch {
        return [80, 80];
      }
    },
  });
}
 
function formatDate(value) {
  if (!value) return new Date().toLocaleDateString("uk-UA");
 
  const d = new Date(value);
 
  if (Number.isNaN(d.getTime())) return value;
 
  return d.toLocaleDateString("uk-UA");
}
 
function formatPrice(value) {
  const number = Number(String(value || "0").replace(/[^\d.,]/g, "").replace(",", "."));
 
  if (Number.isNaN(number)) {
    return value ? String(value) : "";
  }
 
  return Math.round(number).toLocaleString("uk-UA").replace(/\u00A0/g, " ") + " UAH";
}
 
app.post("/generate", async (req, res) => {
  try {
    console.log("🔥 /generate CALLED");
 
    const { clientName, clientDate, managerPhone, managerEmail, managerName, deliveryText, installmentText, longInstallmentText, grandTotalEnabled } = req.body;
 
    const rawProducts = req.body.products || req.body.items || [];
 
    if (!fs.existsSync(TEMPLATE_PATH)) {
      return res.status(500).json({
        error: `Template not found: ${TEMPLATE_PATH}`,
      });
    }
 
    const products = rawProducts.map((p, i) => {
      const title = clean(
        p.title ||
        p.name ||
        p.Name ||
        p["Найменування товару, модель, виробник"]
      );
 
      const description = clean(
        p.description ||
        p.Description
      );
 
      const countryName = clean(
        p.country ||
        p.Country
      );
 
      const photo = clean(
        p.photo ||
        p.Photo
      );
 
      const quantity = Number(p.quantity || p.Quantity || p["Кіль-кість"] || 1);
      const priceNumber = parsePrice(p.price || p.Price || p.priceOriginal);
      const discount = normalizeDiscount(p.discount);
      const priceWithDiscountNumber = getDiscountedPrice(priceNumber, discount);
 
      return {
        index: String(i + 1),
        title,
        description,
        country: countryName ? `Країна виробник: ${countryName}` : "",
        photo,
 
        quantity,
 
        price: formatPrice(priceNumber),
        priceOriginal: formatPrice(priceNumber),
 
        discount,
        showDiscount: discount > 0,
        discountLabel: discount > 0 ? discount + "%" : "",
 
        priceWithDiscount: discount > 0 ? formatPrice(priceWithDiscountNumber) : "",
        finalPrice: formatPrice(priceWithDiscountNumber),
 
        total: formatPrice(priceWithDiscountNumber * quantity),
 
        discountRows: discount > 0
          ? [{
            discountText: "Знижка",
            discountLabel: discount + "%",
            priceWithDiscount: formatPrice(priceWithDiscountNumber)
          }]
          : []
      };
    });
 
    const imageCache = await prefetchImages(products);
    const imageModule = buildImageModule(imageCache);
 
    const zip = loadTemplateZip();
 
    const doc = new Docxtemplater(zip, {
      modules: [imageModule],
      paragraphLoop: true,
      linebreaks: true,
      nullGetter() {
        return "";
      },
    });
 
    let grandTotal = "";
 
    if (grandTotalEnabled) {
      const grandTotalNumber = products.reduce((sum, p) => {
        const numeric = Number(
          String(p.total)
            .replace(/[^\d.,]/g, "")
            .replace(",", ".")
        );
 
        return sum + (isNaN(numeric) ? 0 : numeric);
      }, 0);
 
      grandTotal = `Загальна вартість комерційної пропозиції: ${formatPrice(grandTotalNumber)}`;
    }
 
    doc.render({
      clientName: clientName || "Клієнт",
      clientDate: formatDate(clientDate),
      managerPhone: managerPhone || "",
      managerEmail: managerEmail || "",
      managerName: managerName || "",
      deliveryText: deliveryText || "",
      installmentText: installmentText || "",
      longInstallmentText: longInstallmentText || "",
      products,
      grandTotal,
    });
 
    const buffer = doc.getZip().generate({
      type: "nodebuffer",
      compression: "DEFLATE",
    });
 
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
 
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="KP-demo.docx"'
    );
 
    res.send(buffer);
 
  } catch (err) {
    console.error("❌ FULL ERROR:", err);
 
    const details =
      err.properties?.errors?.map((e) => e.message).join("; ") ||
      err.message;
 
    res.status(500).json({
      error: details,
    });
  }
});
 
const PORT = process.env.PORT || 3001;
 
const server = app.listen(PORT, () => {
  console.log(`🚀 ProposalPro demo server running on http://localhost:${PORT}`);
});
 
server.on("error", (err) => {
  console.error("❌ SERVER ERROR:", err);
});
 