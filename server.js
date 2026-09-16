import express from 'express';
import cors from 'cors';
import multer from 'multer';
import nodemailer from 'nodemailer';
import dotenv from 'dotenv';

dotenv.config({ path: '.env.local' });

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());

app.post('/api/send-email', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file provided' });
    }

    const transporter = nodemailer.createTransport({
      host: 'smtp.office365.com',
      port: 587,
      secure: false,
      auth: {
        user: 'noreply@solutionsandpayroll.com',
        pass: process.env.EMAIL_APP_PASSWORD,
      },
    });

    const htmlContent = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Nuevo Reembolso</title>
<style>
  body { margin:0; padding:0; font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background-color:#f5f5f5; }
  .email-container { max-width:600px; margin:20px auto; background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 2px 12px rgba(0,0,0,0.08); }
  .header { background-color:#1e3a8a; padding:30px 30px 25px 30px; text-align:center; color:#ffffff; position:relative; }
  .logo-container img { height:100px; width:auto; }
  .header::after { content:''; position:absolute; bottom:0; left:0; right:0; height:3px; background-color:#2563eb; }
  .header p { margin:8px 0 0 0; font-size:20px; opacity:0.95; font-weight:400; }
  .content { padding:40px 35px; }
  .greeting { font-size:16px; color:#1e293b; margin-bottom:25px; line-height:1.5; }
  .message-box { background-color:#f8fafc; border:1px solid #e2e8f0; border-left:4px solid #2563eb; padding:24px; margin:25px 0; border-radius:6px; }
  .message-box p { margin:0 0 10px 0; color:#334155; font-size:15px; line-height:1.6; }
  .message-box p:last-child { margin-bottom:0; }
  .highlight { font-weight:600; color:#1e3a8a; }
  .footer { background-color:#f8fafc; padding:30px 35px; text-align:center; border-top:2px solid #e2e8f0; }
  .footer p { margin:8px 0; color:#64748b; font-size:13px; line-height:1.5; }
  .footer-brand { color:#1e3a8a; font-weight:600; font-size:14px; }
</style>
</head>
<body>
<div class="email-container">
  <div class="header">
    <div class="logo-container">
      <img src="https://i.imgur.com/JXCWaXF.png" alt="Solutions &amp; Payroll Logo" />
    </div>
    <p>Sistema de Facturación EOR</p>
  </div>

  <div class="content">
    <div class="greeting">
      Hola,
    </div>

    <div class="message-box">
      <p>
        Le informamos que hay un <span class="highlight">nuevo reembolso</span> disponible.
      </p>
      <p>
        Por favor revise el archivo adjunto para más detalles.
      </p>
    </div>
  </div>

  <div class="footer">
    <p class="footer-brand">Solutions &amp; Payroll</p>
    <p>Este es un mensaje automático, por favor no responder a este correo.</p>
    <p style="margin-top:15px;font-size:12px;color:#94a3b8;">
      © 2026 Solutions &amp; Payroll. Todos los derechos reservados.
    </p>
  </div>
</div>
</body>
</html>`;

    const info = await transporter.sendMail({
      from: '"Solutions & Payroll" <noreply@solutionsandpayroll.com>',
      to: 'automatizacion2@solutionsandpayroll.com',
      subject: 'Nuevo reembolso',
      html: htmlContent,
      attachments: [
        {
          filename: req.file.originalname,
          content: req.file.buffer,
        },
      ],
    });

    res.json({ success: true, messageId: info.messageId });
  } catch (error) {
    console.error('Error sending email:', error);
    res.status(500).json({ error: 'Failed to send email', details: error.message });
  }
});

const PORT = 3001;
app.listen(PORT, () => {
  console.log(`API server running on http://localhost:${PORT}`);
});
