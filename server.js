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

const LOGO_BANNER_PATTERN =
  /(logo|sprite|icon[-_.]|\/icons?\/|favicon|banner|placeholder|avatar|badge|payment|social|footer|header[-_]?bg|watermark|spinner|loading|blank\.gif|whatsapp|instagram|facebook|twitter|youtube|flag[-_])/i;

// Chrome/nav regions repeat on every page of a theme and are the usual
// source of stray "logo PNGs in the background" ending up in the gallery —
// strip them from the DOM before any image scan even runs, rather than
// relying on filename pattern-matching alone to catch every case.
const STRUCTURAL_EXCLUDE_SELECTOR = [
  'header', 'nav', 'footer',
  '.site-header', '.site-footer', '.navbar', '.nav', '.menu', '.main-menu',
  '#masthead', '#colophon',
  '.logo', '.site-logo', '.brand',
  '.widget', 'aside', '.sidebar',
  '.breadcrumb', '.breadcrumbs',
  '.related', '.related-products', '.upsells', '.cross-sells',
  'script', 'style', 'noscript'
].join(', ');

// Dealer/WooCommerce-style product pages almost always wrap the actual car
// photos in one of these gallery widgets, distinct from decorative images
// used elsewhere on the page (theme chrome, category thumbnails, ads).
const GALLERY_CONTAINER_SELECTORS = [
  '.woocommerce-product-gallery',
  '.product-gallery',
  '.product-images',
  '.single-product-images',
  '.images',
  '.car-gallery',
  '.vehicle-gallery',
  '.vehicle-images',
  '.listing-gallery',
  '.listing-images',
  '[class*="product-gallery" i]',
  '[class*="car-gallery" i]',
  '[class*="vehicle-gallery" i]',
  '.swiper-wrapper',
  '.slick-slider',
  '.owl-carousel',
  '[data-lightbox]',
  '[data-fancybox]'
];

function bestImageUrlFromElement($el, pageUrl) {
  const anchorHref = $el.closest('a').attr('href');
  const srcsetLargest = $el
    .attr('srcset')
    ?.split(',')
    .map((s) => s.trim().split(/\s+/)[0])
    .filter(Boolean)
    .pop();

  const candidates = [
    $el.attr('data-large_image'),
    $el.attr('data-large_image_src'),
    $el.attr('data-full'),
    $el.attr('data-zoom-image'),
    anchorHref && /\.(jpe?g|png|webp)(\?|$)/i.test(anchorHref) ? anchorHref : null,
    $el.attr('data-src'),
    $el.attr('data-lazy-src'),
    $el.attr('data-original'),
    srcsetLargest,
    $el.attr('src')
  ].filter(Boolean);

  for (const raw of candidates) {
    const absolute = resolveUrl(pageUrl, raw);
    if (absolute) return absolute;
  }
  return null;
}

function isLikelyPhoto(absoluteUrl, $el) {
  let u;
  try {
    u = new URL(absoluteUrl);
  } catch {
    return false;
  }

  if (/\.svg(\?|$)/i.test(u.pathname)) return false;
  if (LOGO_BANNER_PATTERN.test(u.pathname)) return false;

  const widthAttr = parseInt($el.attr('width') || '0', 10);
  const heightAttr = parseInt($el.attr('height') || '0', 10);
  if ((widthAttr && widthAttr < 150) || (heightAttr && heightAttr < 150)) return false;

  return true;
}

function collectImagesFromScope($, $scope, pageUrl, images) {
  $scope.find('img, source').each((_, el) => {
    const $el = $(el);
    const absolute = bestImageUrlFromElement($el, pageUrl);
    if (absolute && isLikelyPhoto(absolute, $el)) {
      images.add(absolute);
    }
  });
}

function extractCarImages(html, pageUrl) {
  const $ = cheerio.load(html);
  $(STRUCTURAL_EXCLUDE_SELECTOR).remove();

  const images = new Set();

  // Prefer an actual gallery widget over the whole page — this is what
  // keeps unrelated decorative/background images out of the results.
  for (const selector of GALLERY_CONTAINER_SELECTORS) {
    const $container = $(selector);
    if (!$container.length) continue;

    collectImagesFromScope($, $container, pageUrl, images);
    if (images.size) break;
  }

  // No recognizable gallery widget: fall back to the main content area only
  // (never the raw <body>, which is what let theme-wide chrome leak in).
  if (!images.size) {
    const $main = $('main, article, #content, .content, .entry-content, #primary').first();
    collectImagesFromScope($, $main.length ? $main : $('body'), pageUrl, images);
  }

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

const SPEC_TABLE_SELECTOR = [
  'table.woocommerce-product-attributes tr',
  'table.shop_attributes tr',
  '.product-attributes table tr',
  '#tab-additional_information table tr',
  '.woocommerce-Tabs-panel table tr',
  '.car-specs table tr',
  '.vehicle-specs table tr'
].join(', ');

const SPEC_ITEM_SELECTOR = '.spec-item, .vehicle-spec, .car-detail-item, [class*="spec-row" i], [class*="detail-item" i]';

const MAX_SPEC_ENTRIES = 60;

// Pulls every label/value pair the page exposes (WooCommerce "Additional
// Information" tables, <dl> definition lists, generic 2-column tables, and
// label/value widget pairs) into one ordered map, instead of regexing a
// handful of fields out of the flattened body text.
function extractSpecsMap($) {
  const specs = new Map();

  const addEntry = (rawLabel, rawValue) => {
    if (specs.size >= MAX_SPEC_ENTRIES) return;
    const label = (rawLabel || '').replace(/\s+/g, ' ').replace(/:\s*$/, '').trim();
    const value = (rawValue || '').replace(/\s+/g, ' ').trim();
    if (!label || !value || label.length > 60 || value.length > 300) return;
    const key = label.toLowerCase();
    if (!specs.has(key)) specs.set(key, { label, value });
  };

  $(SPEC_TABLE_SELECTOR).each((_, row) => {
    const cells = $(row).find('th, td');
    if (cells.length < 2) return;
    const label = $(cells[0]).text();
    const value = cells
      .slice(1)
      .map((_, c) => $(c).text())
      .get()
      .join(' ');
    addEntry(label, value);
  });

  $('table').each((_, table) => {
    const $table = $(table);
    if ($table.closest(SPEC_TABLE_SELECTOR).length) return; // already handled above
    $table.find('tr').each((_, row) => {
      const cells = $(row).find('th, td');
      if (cells.length !== 2) return; // avoid layout tables with irregular structure
      addEntry($(cells[0]).text(), $(cells[1]).text());
    });
  });

  $('dl').each((_, dl) => {
    const dts = $(dl).find('dt');
    dts.each((_, dt) => {
      const value = $(dt).next('dd').text();
      addEntry($(dt).text(), value);
    });
  });

  $(SPEC_ITEM_SELECTOR).each((_, el) => {
    const $el = $(el);
    const label = $el.find('[class*="label" i]').first().text() || $el.find('span, strong, b').first().text();
    const value = $el.find('[class*="value" i]').first().text();
    if (label && value) addEntry(label, value);
  });

  // Fallback: plain "Label: Value" or "Label - Value" text lines, for themes
  // that don't use a table/dl/label-value widget at all.
  $('li, .spec, .specs, .specification').each((_, el) => {
    if (specs.size >= MAX_SPEC_ENTRIES) return;
    const t = $(el).text().replace(/\s+/g, ' ').trim();
    if (!t || t.length > 150) return;
    const m = t.match(/^([A-Za-z][A-Za-z\s./]{1,40}?)\s*[:\-]\s*(.+)$/);
    if (m) addEntry(m[1], m[2]);
  });

  return specs;
}

function getSpec(specsMap, aliases, avoid = []) {
  for (const alias of aliases) {
    const hit = specsMap.get(alias);
    if (hit) return hit.value;
  }
  // Loose fallback: match a map key that contains the alias as a substring.
  // `avoid` keeps a generic alias (e.g. "color") from grabbing a more
  // specific sibling field's key (e.g. "interior color") when the exact
  // label this field wants isn't present.
  for (const [key, entry] of specsMap) {
    if (avoid.some((a) => key.includes(a))) continue;
    if (aliases.some((alias) => key.includes(alias))) return entry.value;
  }
  return '';
}

function extractDescription($) {
  const selectors = [
    '.woocommerce-product-details__short-description',
    '.product-short-description',
    '#tab-description',
    '.entry-content',
    '.description'
  ];
  for (const sel of selectors) {
    const text = $(sel).first().text().replace(/\s+/g, ' ').trim();
    if (text.length > 20) return text.slice(0, 2000);
  }
  return '';
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

  const specsMap = extractSpecsMap($);

  const year =
    getSpec(specsMap, ['year', 'model year']) ||
    firstMatch(bodyText, [/\b(19[5-9]\d|20[0-4]\d)\b/]) ||
    '';

  const mileage =
    getSpec(specsMap, ['mileage', 'kilometers', 'kilometres', 'km']) ||
    firstMatch(bodyText, [
      /[\d,]{2,}\s*(km|kms|miles|mi)\b/i,
      /mileage\s*[:\-]?\s*[\d,]{2,}\s*(km|miles)?/i
    ]) ||
    '';

  const transmission =
    getSpec(specsMap, ['transmission', 'gearbox']) ||
    firstMatch(bodyText, [/\b(automatic|manual|cvt|tiptronic)\b/i]) ||
    '';

  const fuelType =
    getSpec(specsMap, ['fuel type', 'fuel']) ||
    firstMatch(bodyText, [/\b(petrol|diesel|electric|hybrid|gasoline)\b/i]) ||
    '';

  const bodyType = getSpec(specsMap, ['body type', 'body']);
  const exteriorColor = getSpec(specsMap, ['exterior color', 'exterior colour', 'color', 'colour'], ['interior']);
  const interiorColor = getSpec(specsMap, ['interior color', 'interior colour']);
  const engine = getSpec(specsMap, ['engine capacity', 'engine size', 'engine']);
  const cylinders = getSpec(specsMap, ['cylinders', 'no. of cylinders']);
  const doors = getSpec(specsMap, ['doors', 'no. of doors']);
  const seats = getSpec(specsMap, ['seats', 'seating capacity']);
  const horsepower = getSpec(specsMap, ['horsepower', 'horse power', 'hp']);
  const driveType = getSpec(specsMap, ['drive type', 'drivetrain']);
  const steeringSide = getSpec(specsMap, ['steering side', 'steering']);
  const regionalSpecs = getSpec(specsMap, ['regional specs', 'regional specification', 'specs']);
  const warranty = getSpec(specsMap, ['warranty']);
  const condition = getSpec(specsMap, ['condition']);
  const vin = getSpec(specsMap, ['vin', 'chassis no', 'chassis number']);

  const description = extractDescription($);

  const fullSpecs = Array.from(specsMap.values())
    .map(({ label, value }) => `${label}: ${value}`)
    .join(' | ');

  const images = extractCarImages(html, url);

  return {
    name,
    price,
    year,
    mileage,
    transmission,
    fuelType,
    bodyType,
    exteriorColor,
    interiorColor,
    engine,
    cylinders,
    doors,
    seats,
    horsepower,
    driveType,
    steeringSide,
    regionalSpecs,
    warranty,
    condition,
    vin,
    description,
    fullSpecs,
    url,
    images
  };
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
    { header: 'Body Type', key: 'bodyType', width: 16 },
    { header: 'Exterior Color', key: 'exteriorColor', width: 16 },
    { header: 'Interior Color', key: 'interiorColor', width: 16 },
    { header: 'Engine', key: 'engine', width: 16 },
    { header: 'Cylinders', key: 'cylinders', width: 12 },
    { header: 'Doors', key: 'doors', width: 10 },
    { header: 'Seats', key: 'seats', width: 10 },
    { header: 'Horsepower', key: 'horsepower', width: 14 },
    { header: 'Drive Type', key: 'driveType', width: 14 },
    { header: 'Steering Side', key: 'steeringSide', width: 14 },
    { header: 'Regional Specs', key: 'regionalSpecs', width: 18 },
    { header: 'Warranty', key: 'warranty', width: 18 },
    { header: 'Condition', key: 'condition', width: 14 },
    { header: 'VIN / Chassis No.', key: 'vin', width: 22 },
    { header: 'Description', key: 'description', width: 50 },
    { header: 'Full Specs', key: 'fullSpecs', width: 70 },
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
      bodyType: car.bodyType,
      exteriorColor: car.exteriorColor,
      interiorColor: car.interiorColor,
      engine: car.engine,
      cylinders: car.cylinders,
      doors: car.doors,
      seats: car.seats,
      horsepower: car.horsepower,
      driveType: car.driveType,
      steeringSide: car.steeringSide,
      regionalSpecs: car.regionalSpecs,
      warranty: car.warranty,
      condition: car.condition,
      vin: car.vin,
      description: car.description,
      fullSpecs: car.fullSpecs,
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

  sheet.autoFilter = { from: 'A1', to: 'X1' };
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
      bodyType: details.bodyType,
      exteriorColor: details.exteriorColor,
      interiorColor: details.interiorColor,
      engine: details.engine,
      cylinders: details.cylinders,
      doors: details.doors,
      seats: details.seats,
      horsepower: details.horsepower,
      driveType: details.driveType,
      steeringSide: details.steeringSide,
      regionalSpecs: details.regionalSpecs,
      warranty: details.warranty,
      condition: details.condition,
      vin: details.vin,
      description: details.description,
      fullSpecs: details.fullSpecs,
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
