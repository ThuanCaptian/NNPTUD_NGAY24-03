var express = require("express");
var router = express.Router();
let userModel = require("../schemas/users");
let roleModel = require("../schemas/roles");
let { validatedResult, CreateAnUserValidator, ModifyAnUserValidator } = require('../utils/validator')
let userController = require('../controllers/users')
let { CheckLogin, checkRole } = require('../utils/authHandler')
let { uploadExcel } = require('../utils/uploadHandler')
let { sendUserPasswordMail } = require('../utils/sendMail')
let exceljs = require('exceljs')
let crypto = require('crypto')
let path = require('path')
let fs = require('fs')

function getCellStringValue(cell) {
  if (!cell) return "";
  let value = cell.value;
  if (value === null || value === undefined) return "";

  if (typeof value === "object") {
    if (value.result !== undefined && value.result !== null) value = value.result;
    else if (value.text) value = value.text;
    else if (Array.isArray(value.richText)) value = value.richText.map(e => e.text).join("");
    else if (value.hyperlink) value = value.hyperlink;
  }
  return String(value).trim();
}

function generatePassword(length = 16) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+-=[]{}";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars[crypto.randomInt(0, chars.length)];
  }
  return result;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function sendPasswordMailWithRetry(email, username, password) {
  let lastError = null;
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await sendUserPasswordMail(email, username, password);
      return;
    } catch (error) {
      lastError = error;
      const isRateLimited = error.message && error.message.includes("Too many emails per second");
      if (isRateLimited && attempt < maxAttempts) {
        await sleep(1200 * attempt);
        continue;
      }
      break;
    }
  }
  throw lastError;
}


router.get("/", CheckLogin, checkRole("ADMIN","MODERATOR"), async function (req, res, next) {//ADMIN
  let users = await userController.GetAllUser()
  res.send(users);
});

router.get("/:id", async function (req, res, next) {
  let result = await userController.GetUserById(
    req.params.id
  )
  if (result) {
    res.send(result);
  } else {
    res.status(404).send({ message: "id not found" })
  }
});

router.post("/", CreateAnUserValidator, validatedResult, async function (req, res, next) {
  
  try {
    let user = await userController.CreateAnUser(
      req.body.username, req.body.password,
      req.body.email, req.body.role
    )
    res.send(user);
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

router.post("/import/excel", uploadExcel.single('file'), async function (req, res, next) {
  if (!req.file) {
    return res.status(404).send({ message: "file not found" });
  }

  let pathFile = path.join(__dirname, "../uploads", req.file.filename);
  try {
    let roleUser = await roleModel.findOne({
      isDeleted: false,
      name: new RegExp("^user$", "i")
    });
    if (!roleUser) {
      roleUser = new roleModel({
        name: "USER",
        description: "Default role for imported users"
      });
      await roleUser.save();
    }

    let workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(pathFile);
    let worksheet = workbook.worksheets[0];
    if (!worksheet || worksheet.rowCount < 2) {
      return res.status(400).send({ message: "file excel khong co du lieu" });
    }

    let imported = 0;
    let failed = 0;
    let emailed = 0;
    let errors = [];

    for (let index = 2; index <= worksheet.rowCount; index++) {
      let row = worksheet.getRow(index);
      let username = getCellStringValue(row.getCell(1));
      let email = getCellStringValue(row.getCell(2)).toLowerCase();

      if (!username && !email) {
        continue;
      }
      if (!username || !email) {
        failed++;
        errors.push({
          row: index,
          message: "thieu username hoac email"
        });
        continue;
      }

      let randomPassword = generatePassword(16);
      try {
        let newUser = await userController.CreateAnUser(
          username,
          randomPassword,
          email,
          roleUser._id
        );
        imported++;

        try {
          await sendPasswordMailWithRetry(newUser.email, newUser.username, randomPassword);
          emailed++;
        } catch (mailError) {
          errors.push({
            row: index,
            message: "tao user thanh cong nhung gui email that bai: " + mailError.message
          });
        }
      } catch (error) {
        failed++;
        errors.push({
          row: index,
          message: error.message
        });
      }
    }

    res.send({
      message: "import user thanh cong",
      totalRows: worksheet.rowCount - 1,
      imported: imported,
      failed: failed,
      emailed: emailed,
      errors: errors
    });
  } catch (err) {
    res.status(400).send({ message: err.message });
  } finally {
    if (fs.existsSync(pathFile)) {
      fs.unlinkSync(pathFile);
    }
  }
});

router.put("/:id", ModifyAnUserValidator, validatedResult, async function (req, res, next) {
  try {
    let id = req.params.id;
    let updatedItem = await userModel.findByIdAndUpdate
      (id, req.body, { new: true });

    if (!updatedItem) return res.status(404).send({ message: "id not found" });

    let populated = await userModel
      .findById(updatedItem._id)
    res.send(populated);
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

router.delete("/:id", async function (req, res, next) {
  try {
    let id = req.params.id;
    let updatedItem = await userModel.findByIdAndUpdate(
      id,
      { isDeleted: true },
      { new: true }
    );
    if (!updatedItem) {
      return res.status(404).send({ message: "id not found" });
    }
    res.send(updatedItem);
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

module.exports = router;
