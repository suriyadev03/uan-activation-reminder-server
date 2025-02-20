require('dotenv').config();
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const cors = require('cors');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors()); // Enable CORS if needed

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

    const total = data.length;
    let sentCount = 0;

    for (const row of data) {
      const email = row['Mail'];
      const estName = row['office Name'];
      const action = row['Unactivate'];

      if (!email) continue;

      try {
        await transporter.sendMail({
          from: `"EPFO Office" <${process.env.EMAIL_USER}>`,
          to: email,
          subject: `Reminder: UAN Activation Pending for ${estName} Office`,
          text: `Dear ${estName} Office HR,\n\nThis mail is regarding ${estName} Office. We noticed that ${action} employees have not activated their UAN.\n\nThis is an reminder for the HR team to ensure UAN activation is completed for all employees.\n\nRegards,\nEPFO Office`,
          html: `
            <p>Dear <strong>${estName} Office HR</strong>,</p>
            <p>This mail is regarding <strong>${estName} Office</strong>. We noticed that <strong>${action}</strong> employees have not activated their UAN.</p>
            <p>This is an <strong>reminder</strong> for the <strong>HR team</strong> to ensure UAN activation is completed for all employees.</p>
            <p>Regards,<br><strong>EPFO Office</strong></p>
          `,
        });

        sentCount++;
      } catch (error) {
        console.error(`Failed to send email to ${email}:`, error);
      }
    }

    res.json({ sent: sentCount, total });
  } catch (error) {
    console.error('File processing error:', error);
    res.status(500).send('Error processing file or sending emails');
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
