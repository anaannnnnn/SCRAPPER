const path = require('path');
const fs = require('fs');
const fsp = fs.promises;
const crypto = require('crypto');
const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const archiver = require('archiver');

const app = express();
const PORT = process.env.PORT || 3000;

const TEMP_DIR = path.join(__dirname, 'temp');
const DOWNLOADS_DIR = path.join(__dirname, 'downloads');

for (const dir of [TEMP_DIR, DOWNLOADS_DIR]) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use('/downloads', express.static(DOWNLOADS_DIR));

const HTTP_HEADERS = {
  'User-Agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
};

const REQUEST_TIMEOUT_MS = 15000;
const MAX_LISTINGS = 200;
const MAX_IMAGES_PER_CAR = 15;

// ---------- helpers ----------

function makeLogger() {
  const logs = [];
  return {
    logs,
    push(message, type = 'info') {
      logs.push({ message, type });
    }
  };
}

function sanitizeFolderName(name) {
  return (name || 'unnamed-car')
    .toString()
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80) || 'unnamed-car';
}

function resolveUrl(base, href) {
  try {
    return new URL(href, base).toString();
  } catch {
    return null;
  }
}

async function fetchHtml(url, attempts = 3) {
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      const res = await axios.get(url, {
        headers: HTTP_HEADERS,
        timeout: REQUEST_TIMEOUT_MS,
        responseType: 'text',
        validateStatus: (s) => s >= 200 && s < 400
      });
      return res.data;
    } catch (err) {
      lastErr = err;
      // Only worth retrying on transient network/timeout errors, not on a
      // real 4xx/5xx response (validateStatus already filters those into
      // axios's own error, but we still don't want to hammer a hard failure).
      const transient = !err.response || err.code === 'ECONNABORTED' || err.code === 'ETIMEDOUT' || err.code === 'ECONNRESET';
      if (i < attempts && transient) {
        await new Promise((r) => setTimeout(r, 750 * i));
        continue;
      }
      break;
    }
  }
  throw lastErr;
}

// Runs async tasks with a concurrency cap instead of one-at-a-time or all-at-once.
async function runWithConcurrency(items, limit, worker) {
  const results = new Array(items.length);
  let nextIndex = 0;
  async function runNext() {
    const i = nextIndex++;
    if (i >= items.length) return;
    results[i] = await worker(items[i], i);
    await runNext();
  }
  const runners = Array.from({ length: Math.min(limit, items.length) }, runNext);
  await Promise.all(runners);
  return results;
}

function buildPaginationUrl(baseUrl, pageNum) {
  if (pageNum <= 1) return baseUrl;
  const url = new URL(baseUrl);
  const trimmedPath = url.pathname.replace(/\/+$/, '');
  url.pathname = `${trimmedPath}/page/${pageNum}/`;
  return url.toString();
}

function buildPaginationUrlAlt(baseUrl, pageNum) {
  const url = new URL(baseUrl);
  url.searchParams.set('page', String(pageNum));
  return url.toString();
}

const NAV_LINK_PATTERN = /(login|signup|sign-up|register|cart|checkout|wishlist|compare|contact|about|privacy|terms|blog|faq|careers|sitemap|\.pdf$|\.jpg$|\.png$|mailto:|tel:|javascript:|#)/i;

function extractListingLinks(html, pageUrl) {
  const $ = cheerio.load(html);
  const origin = new URL(pageUrl).origin;
  const basePathSegments = new URL(pageUrl).pathname.split('/').filter(Boolean);

  const candidates = new Map();

  $('a[href]').each((_, el) => {
    const href = $(el).attr('href');
    if (!href) return;

    const absolute = resolveUrl(pageUrl, href);
    if (!absolute) return;

    let u;
    try {
      u = new URL(absolute);
    } catch {
      return;
    }

    if (u.origin !== origin) return;
    if (NAV_LINK_PATTERN.test(u.pathname) || NAV_LINK_PATTERN.test(href)) return;
    if (/\/page\/\d+/i.test(u.pathname) || /[?&]page=\d+/i.test(u.search) || /[?&]paged=\d+/i.test(u.search)) return;

    const segments = u.pathname.split('/').filter(Boolean);
    if (segments.length <= basePathSegments.length) return;

    const contextText = ($(el).closest('article, li, div, [class]').attr('class') || '') + ' ' + ($(el).attr('class') || '');
    const isLikelyListing =
      /car|vehicle|listing|product|item|auto/i.test(contextText) ||
      /car|vehicle|listing|product|item|auto/i.test(u.pathname);

    const key = u.origin + u.pathname;
    if (!candidates.has(key) && (isLikelyListing || segments.length >= basePathSegments.length + 1)) {
      candidates.set(key, key);
    }
  });

  return Array.from(candidates.values());
}

const LOGO_BANNER_PATTERN = /(logo|sprite|icon|banner|placeholder|avatar|badge|payment|social|footer|header-bg|watermark|spinner|loading|blank\.gif)/i;

function extractCarImages(html, pageUrl) {
  const $ = cheerio.load(html);
  const images = new Set();

  $('img, source').each((_, el) => {
    const $el = $(el);
    const candidates = [
      $el.attr('src'),
      $el.attr('data-src'),
      $el.attr('data-lazy-src'),
      $el.attr('data-original'),
      $el.attr('srcset')?.split(',').pop()?.trim().split(' ')[0]
    ].filter(Boolean);

    for (const raw of candidates) {
      const absolute = resolveUrl(pageUrl, raw);
      if (!absolute) continue;

      let u;
      try {
        u = new URL(absolute);
      } catch {
        continue;
      }

      if (/\.svg(\?|$)/i.test(u.pathname)) continue;
      if (LOGO_BANNER_PATTERN.test(u.pathname)) continue;

      const widthAttr = parseInt($el.attr('width') || '0', 10);
      const heightAttr = parseInt($el.attr('height') || '0', 10);
      if ((widthAttr && widthAttr < 120) || (heightAttr && heightAttr < 120)) continue;

      images.add(absolute);
    }
  });

  return Array.from(images).slice(0, MAX_IMAGES_PER_CAR);
}

function firstMatch(text, patterns) {
  for (const p of patterns) {
    const m = text.match(p);
    // Always take the full match, not a capture group: several of the
    // patterns below use a group only to anchor on a currency/unit token,
    // and the group alone (e.g. just "AED") throws away the actual value.
    if (m) return m[0].trim();
  }
  return '';
}

// Many dealer-site themes reuse a generic archive/shop heading ("Cars",
// "Shop", "Inventory", "All Vehicles"...) as the <h1> on every listing page,
// including individual car pages, instead of the car's own title. Detect
// that so we can fall back to a per-page title source instead.
const GENERIC_HEADING_PATTERN = /^(cars?|vehicles?|shop|products?|inventory|listings?|home|catalog|our\s+(cars?|vehicles?|inventory))$/i;

function cleanTitleTag(title) {
  // Strip a trailing " | Site Name" / " - Site Name" suffix some themes add.
  return title.replace(/\s*[|–—-]\s*[^|–—-]{1,40}$/, '').trim() || title.trim();
}

function extractCarDetails(html, url) {
  const $ = cheerio.load(html);
  const bodyText = $('body').text().replace(/\s+/g, ' ');

  const ogTitle = $('meta[property="og:title"]').attr('content')?.trim();
  const h1Text = $('h1').first().text().trim();
  const titleTag = $('title').text().trim();

  const name =
    (ogTitle && !GENERIC_HEADING_PATTERN.test(ogTitle) && ogTitle) ||
    (h1Text && !GENERIC_HEADING_PATTERN.test(h1Text) && h1Text) ||
    (titleTag && cleanTitleTag(titleTag)) ||
    h1Text ||
    'Unknown Car';

  // Prefer a real price widget when the theme exposes one (WooCommerce and
  // most storefront themes do) — far more reliable than scanning body text.
  const priceSelectorText = $('.price, .woocommerce-Price-amount, [class*="price" i]')
    .first()
    .text()
    .replace(/\s+/g, ' ')
    .trim();
  const priceLooksValid = /\d/.test(priceSelectorText) && priceSelectorText.length < 40;

  const price = (priceLooksValid && priceSelectorText) || firstMatch(bodyText, [
    /(AED|USD|EUR|GBP|\$|price)\s*[:\-]?\s*[\d,]{3,}/i,
    /[\d,]{4,}\s*(AED|USD|EUR|GBP)/i
  ]) || '';

  const year = firstMatch(bodyText, [/\b(19[5-9]\d|20[0-4]\d)\b/]) || '';

  const mileage = firstMatch(bodyText, [
    /[\d,]{2,}\s*(km|kms|miles|mi)\b/i,
    /mileage\s*[:\-]?\s*[\d,]{2,}\s*(km|miles)?/i
  ]) || '';

  const transmission = firstMatch(bodyText, [/\b(automatic|manual|cvt|tiptronic)\b/i]) || '';

  const fuelType = firstMatch(bodyText, [/\b(petrol|diesel|electric|hybrid|gasoline)\b/i]) || '';

  const specParts = [];
  $('li, table tr, .spec, .specs, .specification').each((_, el) => {
    const t = $(el).text().replace(/\s+/g, ' ').trim();
    if (t && t.length < 120 && /[:\-]/.test(t)) specParts.push(t);
  });
  const specs = Array.from(new Set(specParts)).slice(0, 15).join(' | ');

  const images = extractCarImages(html, url);

  return { name, price, year, mileage, transmission, fuelType, specs, url, images };
}

async function downloadImage(imageUrl, destPath) {
  const res = await axios.get(imageUrl, {
    headers: HTTP_HEADERS,
    timeout: REQUEST_TIMEOUT_MS,
    responseType: 'arraybuffer',
    validateStatus: (s) => s >= 200 && s < 400
  });
  await fsp.writeFile(destPath, res.data);
}

function imageExtensionFromUrl(url) {
  const match = url.split('?')[0].match(/\.(jpg|jpeg|png|webp|gif)$/i);
  return match ? match[1].toLowerCase() : 'jpg';
}

async function buildExcelWorkbook(cars, outputPath) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Universal Car Inventory Scraper';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Car Inventory');

  sheet.columns = [
    { header: 'Car Name', key: 'name', width: 40 },
    { header: 'Price', key: 'price', width: 18 },
    { header: 'Year', key: 'year', width: 10 },
    { header: 'Mileage', key: 'mileage', width: 16 },
    { header: 'Transmission', key: 'transmission', width: 16 },
    { header: 'Fuel Type', key: 'fuelType', width: 14 },
    { header: 'Specs', key: 'specs', width: 60 },
    { header: 'Image Count', key: 'imageCount', width: 14 },
    { header: 'Listing URL', key: 'url', width: 50 }
  ];

  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E3A5F' }
    };
    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FF0B0F14' } } };
  });
  headerRow.height = 22;

  cars.forEach((car, idx) => {
    const row = sheet.addRow({
      name: car.name,
      price: car.price,
      year: car.year,
      mileage: car.mileage,
      transmission: car.transmission,
      fuelType: car.fuelType,
      specs: car.specs,
      imageCount: car.imageCount,
      url: car.url
    });
    if (idx % 2 === 1) {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F6FA' } };
      });
    }
    row.eachCell((cell) => {
      cell.alignment = { vertical: 'middle', wrapText: false };
      cell.border = { bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } } };
    });
  });

  sheet.autoFilter = { from: 'A1', to: 'I1' };
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  await workbook.xlsx.writeFile(outputPath);
}

async function zipDirectory(sourceDir, zipPath) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(zipPath);
    const archive = archiver('zip', { zlib: { level: 9 } });

    output.on('close', resolve);
    archive.on('error', reject);

    archive.pipe(output);
    archive.directory(sourceDir, false);
    archive.finalize();
  });
}

async function removeDirSafe(dir) {
  try {
    await fsp.rm(dir, { recursive: true, force: true });
  } catch {
    // best-effort cleanup
  }
}

// ---------- main scrape pipeline ----------

async function scrapeInventory({ targetUrl, maxPages }, logger) {
  const jobId = crypto.randomBytes(6).toString('hex');
  const workDir = path.join(TEMP_DIR, jobId);
  const imagesDir = path.join(workDir, 'images');
  await fsp.mkdir(imagesDir, { recursive: true });

  logger.push(`Job ${jobId} started`, 'info');

  const allLinks = new Set();
  let pagesScanned = 0;

  for (let page = 1; page <= maxPages; page++) {
    const pageUrl = page === 1 ? targetUrl : buildPaginationUrl(targetUrl, page);
    logger.push(`Fetching listing page ${page}: ${pageUrl}`, 'info');

    let html;
    try {
      html = await fetchHtml(pageUrl);
    } catch (err) {
      if (page > 1) {
        try {
          html = await fetchHtml(buildPaginationUrlAlt(targetUrl, page));
        } catch {
          logger.push(`Page ${page} unreachable, stopping pagination.`, 'warn');
          break;
        }
      } else {
        throw new Error(`Could not reach target URL: ${err.message}`);
      }
    }

    pagesScanned++;
    const links = extractListingLinks(html, pageUrl);
    const newLinks = links.filter((l) => !allLinks.has(l));
    newLinks.forEach((l) => allLinks.add(l));

    logger.push(`Page ${page}: found ${links.length} listing links (${newLinks.length} new)`, 'ok');

    if (newLinks.length === 0 && page > 1) {
      logger.push('No new listings found, stopping pagination early.', 'warn');
      break;
    }
    if (allLinks.size >= MAX_LISTINGS) {
      logger.push(`Reached listing cap (${MAX_LISTINGS}), stopping pagination.`, 'warn');
      break;
    }
  }

  const listingUrls = Array.from(allLinks).slice(0, MAX_LISTINGS);
  logger.push(`Total unique car listings discovered: ${listingUrls.length}`, 'accent');

  const cars = [];
  let totalImages = 0;
  const usedFolderNames = new Set();

  const LISTING_CONCURRENCY = 4;
  const IMAGE_CONCURRENCY = 5;

  // Fetch + parse every listing page concurrently (bounded) instead of one
  // at a time — for a real inventory (dozens to hundreds of cars) a fully
  // sequential loop can take long enough to run past a hosting platform's
  // own request timeout even though every individual fetch succeeds.
  const detailResults = await runWithConcurrency(listingUrls, LISTING_CONCURRENCY, async (url, i) => {
    logger.push(`[${i + 1}/${listingUrls.length}] Extracting: ${url}`, 'info');
    try {
      const html = await fetchHtml(url);
      return { url, details: extractCarDetails(html, url) };
    } catch (err) {
      logger.push(`Failed to fetch listing: ${err.message}`, 'warn');
      return { url, details: null };
    }
  });

  for (const { details } of detailResults) {
    if (!details) continue;

    let folderName = sanitizeFolderName(details.name);
    let uniqueFolder = folderName;
    let dupCount = 1;
    while (usedFolderNames.has(uniqueFolder)) {
      dupCount++;
      uniqueFolder = `${folderName} (${dupCount})`;
    }
    usedFolderNames.add(uniqueFolder);

    const carImagesDir = path.join(imagesDir, uniqueFolder);
    await fsp.mkdir(carImagesDir, { recursive: true });

    const downloadFlags = await runWithConcurrency(details.images, IMAGE_CONCURRENCY, async (imgUrl, imgIdx) => {
      const ext = imageExtensionFromUrl(imgUrl);
      const destPath = path.join(carImagesDir, `image_${imgIdx + 1}.${ext}`);
      try {
        await downloadImage(imgUrl, destPath);
        return true;
      } catch {
        return false; // skip broken image
      }
    });
    const downloadedCount = downloadFlags.filter(Boolean).length;

    totalImages += downloadedCount;
    logger.push(`  -> ${details.name} | ${downloadedCount} image(s) saved`, downloadedCount ? 'ok' : 'warn');

    cars.push({
      name: details.name,
      price: details.price,
      year: details.year,
      mileage: details.mileage,
      transmission: details.transmission,
      fuelType: details.fuelType,
      specs: details.specs,
      url: details.url,
      imageCount: downloadedCount
    });
  }

  logger.push('Generating styled Excel spreadsheet...', 'info');
  const excelPath = path.join(workDir, 'Car_Inventory_Database.xlsx');
  await buildExcelWorkbook(cars, excelPath);
  logger.push('Excel spreadsheet created.', 'ok');

  logger.push('Bundling images and spreadsheet into ZIP archive...', 'info');
  const zipFileName = `Car_Inventory_${jobId}.zip`;
  const zipPath = path.join(DOWNLOADS_DIR, zipFileName);
  await zipDirectory(workDir, zipPath);
  logger.push('ZIP archive ready.', 'ok');

  await removeDirSafe(workDir);
  logger.push('Temporary working files cleaned up.', 'info');

  return {
    cars,
    stats: {
      pagesScanned,
      carsFound: cars.length,
      imagesDownloaded: totalImages
    },
    downloadUrl: `/downloads/${zipFileName}`,
    fileName: zipFileName
  };
}

// ---------- routes ----------

app.post('/api/scrape', async (req, res) => {
  const logger = makeLogger();
  try {
    const { targetUrl, maxPages } = req.body || {};

    if (!targetUrl || typeof targetUrl !== 'string') {
      return res.status(400).json({ success: false, error: 'targetUrl is required.' });
    }

    let parsedUrl;
    try {
      parsedUrl = new URL(targetUrl);
    } catch {
      return res.status(400).json({ success: false, error: 'targetUrl must be a valid URL.' });
    }
    if (!/^https?:$/.test(parsedUrl.protocol)) {
      return res.status(400).json({ success: false, error: 'targetUrl must use http or https.' });
    }

    const safeMaxPages = Math.min(Math.max(parseInt(maxPages, 10) || 1, 1), 50);

    const result = await scrapeInventory({ targetUrl, maxPages: safeMaxPages }, logger);

    res.json({
      success: true,
      logs: logger.logs,
      cars: result.cars,
      stats: result.stats,
      downloadUrl: result.downloadUrl,
      fileName: result.fileName
    });
  } catch (err) {
    logger.push(`Fatal error: ${err.message}`, 'err');
    res.status(500).json({ success: false, error: err.message, logs: logger.logs });
  }
});

app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok' });
});

app.listen(PORT, () => {
  console.log(`Universal Car Inventory Scraper running on http://localhost:${PORT}`);
});
