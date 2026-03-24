const nodemailer = require("nodemailer");

const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || "sandbox.smtp.mailtrap.io",
    port: Number(process.env.SMTP_PORT || 2525),
    secure: false,
    auth: {
        user: process.env.SMTP_USER || "689dc18e08ae57",
        pass: process.env.SMTP_PASS || "f0bf2c0324194a",
    },
});

module.exports = {
    sendMail: async function (to, url) {
        await transporter.sendMail({
            from: 'admin@haha.com',
            to: to,
            subject: "reset password email",
            text: "click vao day de doi pass", // Plain-text version of the message
            html: "click vao <a href=" + url+ ">day</a> de doi pass", // HTML version of the message
        })
    },
    sendUserPasswordMail: async function (to, username, password) {
        await transporter.sendMail({
            from: "admin@haha.com",
            to: to,
            subject: "Thong tin tai khoan moi",
            text: "Tai khoan: " + username + " | Mat khau: " + password,
            html:
                "<p>Chao " + username + ",</p>" +
                "<p>Tai khoan cua ban da duoc tao thanh cong.</p>" +
                "<p><b>Username:</b> " + username + "</p>" +
                "<p><b>Password:</b> " + password + "</p>" +
                "<p>Ban nen dang nhap va doi mat khau sau lan dau su dung.</p>",
        });
    }
}

// Send an email using async/await
