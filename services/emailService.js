const nodemailer = require('nodemailer');

class EmailService {
  constructor() {
    this.transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST || 'smtp.gmail.com',
      port: Number(process.env.SMTP_PORT || 587),
      secure: String(process.env.SMTP_SECURE || 'false').toLowerCase() === 'true',
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASSWORD
      }
    });
  }

  async send({ to, subject, text, attachments = [] }) {
    if (!to) throw new Error('No hay destinatario para la notificación');
    if (!process.env.SMTP_USER || !process.env.SMTP_PASSWORD) {
      throw new Error('Faltan SMTP_USER o SMTP_PASSWORD');
    }

    return this.transporter.sendMail({
      from: process.env.SMTP_FROM || process.env.SMTP_USER,
      to,
      subject,
      text,
      attachments
    });
  }
}

module.exports = new EmailService();
