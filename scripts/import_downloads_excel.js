const path = require("path");
const crypto = require("crypto");
const mongoose = require("mongoose");
const exceljs = require("exceljs");
const slugify = require("slugify");

const userModel = require("../schemas/users");
const roleModel = require("../schemas/roles");
const categoryModel = require("../schemas/categories");
const productModel = require("../schemas/products");
const inventoryModel = require("../schemas/inventories");
const { sendUserPasswordMail } = require("../utils/sendMail");

const MONGO_URI = process.env.MONGO_URI || "mongodb://localhost:27017/NNPTUD-C3";
const USER_FILE = process.env.USER_XLSX || "C:/Users/Admin/Downloads/user.xlsx";
const PRODUCT_FILE =
  process.env.PRODUCT_XLSX || "C:/Users/Admin/Downloads/products_3000_rows.xlsx";

function getCellStringValue(cell) {
  if (!cell) return "";
  let value = cell.value;
  if (value === null || value === undefined) return "";

  if (typeof value === "object") {
    if (value.result !== undefined && value.result !== null) value = value.result;
    else if (value.text) value = value.text;
    else if (Array.isArray(value.richText)) value = value.richText.map((e) => e.text).join("");
    else if (value.hyperlink) value = value.hyperlink;
  }
  return String(value).trim();
}

function getCellNumberValue(cell) {
  const value = getCellStringValue(cell);
  if (!value) return 0;
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return 0;
  return parsed;
}

function generatePassword(length = 16) {
  const chars =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars[crypto.randomInt(0, chars.length)];
  }
  return result;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function sendPasswordMailWithRetry(email, username, password) {
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      await sendUserPasswordMail(email, username, password);
      return true;
    } catch (error) {
      lastError = error;
      const isRateLimited =
        error.message && error.message.includes("Too many emails per second");
      if (isRateLimited && attempt < 3) {
        await sleep(1200 * attempt);
        continue;
      }
      break;
    }
  }
  throw lastError;
}

function buildUniqueSlug(baseValue, fallback) {
  const slug = slugify(baseValue || fallback, {
    replacement: "-",
    lower: true,
    strict: true,
    trim: true,
  });
  return slug || fallback;
}

async function ensureUserRole() {
  let roleUser = await roleModel.findOne({
    isDeleted: false,
    name: new RegExp("^user$", "i"),
  });
  if (!roleUser) {
    roleUser = new roleModel({
      name: "USER",
      description: "Default role for imported users",
    });
    await roleUser.save();
  }
  return roleUser;
}

async function getOrCreateCategory(categoryName, categoryCache) {
  const cacheKey = categoryName.toLowerCase();
  if (categoryCache.has(cacheKey)) {
    return categoryCache.get(cacheKey);
  }

  let category = await categoryModel.findOne({
    isDeleted: false,
    name: new RegExp("^" + categoryName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "$", "i"),
  });
  if (!category) {
    category = new categoryModel({
      name: categoryName,
      slug: buildUniqueSlug(categoryName, "category"),
    });
    try {
      await category.save();
    } catch (error) {
      if (error.code === 11000) {
        category = await categoryModel.findOne({
          name: new RegExp("^" + categoryName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "$", "i"),
        });
      } else {
        throw error;
      }
    }
  }

  categoryCache.set(cacheKey, category);
  return category;
}

async function importUsers(userFile) {
  const workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(userFile);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("Khong tim thay worksheet trong file user");
  }

  const roleUser = await ensureUserRole();
  const summary = {
    totalRows: Math.max(worksheet.rowCount - 1, 0),
    created: 0,
    skipped: 0,
    emailed: 0,
    mailFailed: 0,
    errors: [],
  };

  for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
    const row = worksheet.getRow(rowIndex);
    const username = getCellStringValue(row.getCell(1));
    const email = getCellStringValue(row.getCell(2)).toLowerCase();

    if (!username && !email) continue;
    if (!username || !email) {
      summary.skipped++;
      summary.errors.push({ row: rowIndex, message: "thieu username hoac email" });
      continue;
    }

    const existed = await userModel.findOne({
      $or: [{ username: username }, { email: email }],
    });
    if (existed) {
      summary.skipped++;
      continue;
    }

    const rawPassword = generatePassword(16);
    const newUser = new userModel({
      username: username,
      password: rawPassword,
      email: email,
      role: roleUser._id,
    });

    try {
      await newUser.save();
      summary.created++;
    } catch (error) {
      summary.skipped++;
      summary.errors.push({ row: rowIndex, message: error.message });
      continue;
    }

    try {
      await sendPasswordMailWithRetry(newUser.email, newUser.username, rawPassword);
      summary.emailed++;
      await sleep(350);
    } catch (mailError) {
      summary.mailFailed++;
      summary.errors.push({
        row: rowIndex,
        message: "tao user thanh cong nhung gui email that bai: " + mailError.message,
      });
    }
  }

  return summary;
}

async function importProducts(productFile) {
  const workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(productFile);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("Khong tim thay worksheet trong file product");
  }

  const summary = {
    totalRows: Math.max(worksheet.rowCount - 1, 0),
    createdProducts: 0,
    updatedProducts: 0,
    createdInventories: 0,
    updatedInventories: 0,
    skipped: 0,
    errors: [],
  };

  const categoryCache = new Map();

  for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
    const row = worksheet.getRow(rowIndex);
    const title = getCellStringValue(row.getCell(2));
    const categoryName = getCellStringValue(row.getCell(3));
    const price = getCellNumberValue(row.getCell(4));
    const stock = getCellNumberValue(row.getCell(5));

    if (!title && !categoryName) continue;
    if (!title || !categoryName) {
      summary.skipped++;
      summary.errors.push({ row: rowIndex, message: "thieu title hoac category" });
      continue;
    }

    try {
      const category = await getOrCreateCategory(categoryName, categoryCache);
      let product = await productModel.findOne({ title: title });
      if (!product) {
        product = new productModel({
          title: title,
          slug: buildUniqueSlug(title, "product"),
          price: price,
          category: category._id,
        });
        await product.save();
        summary.createdProducts++;
      } else {
        product.price = price;
        product.category = category._id;
        await product.save();
        summary.updatedProducts++;
      }

      const inventory = await inventoryModel.findOne({ product: product._id });
      if (!inventory) {
        const newInventory = new inventoryModel({
          product: product._id,
          stock: stock,
        });
        await newInventory.save();
        summary.createdInventories++;
      } else {
        inventory.stock = stock;
        await inventory.save();
        summary.updatedInventories++;
      }
    } catch (error) {
      summary.skipped++;
      summary.errors.push({ row: rowIndex, message: error.message });
    }
  }

  return summary;
}

async function run() {
  const userFile = path.resolve(USER_FILE);
  const productFile = path.resolve(PRODUCT_FILE);

  console.log("MONGO_URI:", MONGO_URI);
  console.log("USER_FILE:", userFile);
  console.log("PRODUCT_FILE:", productFile);

  await mongoose.connect(MONGO_URI);
  try {
    const userSummary = await importUsers(userFile);
    const productSummary = await importProducts(productFile);

    console.log("=== USER IMPORT SUMMARY ===");
    console.log(JSON.stringify(userSummary, null, 2));
    console.log("=== PRODUCT IMPORT SUMMARY ===");
    console.log(JSON.stringify(productSummary, null, 2));
  } finally {
    await mongoose.disconnect();
  }
}

run().catch((error) => {
  console.error("IMPORT_FAILED:", error.message);
  process.exit(1);
});
