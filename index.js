import express from "express";
import { dirname } from "path";
import { fileURLToPath } from "url";
import bodyParser from "body-parser";
import nodemailer from "nodemailer";
import qrcode from "qrcode";
import fs from "fs";
import { promisify } from "util";
import { createCanvas, loadImage } from "canvas";
import { v4 as uuidv4 } from 'uuid';
import mongoose from 'mongoose'; 
import XLSX from 'xlsx';
import Registration from './models/Registration.js'; // Import the model
import { createObjectCsvWriter } from 'csv-writer';

const __dirname = dirname(fileURLToPath(import.meta.url));

const port = 3001;
const app = express();

app.use(bodyParser.urlencoded({ extended: true }));

const writeFile = promisify(fs.writeFile);

app.use(express.static('static'));

// MongoDB connection
mongoose.connect('mongodb+srv://bhattacharjeeagnivajobs:MgCW8rI2JIuaDJ0I@cluster0.pzt6xl6.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0', {
  useNewUrlParser: true,
  useUnifiedTopology: true
}).then(() => {
  console.log('Connected to MongoDB');
}).catch(err => {
  console.error('Error connecting to MongoDB', err);
});

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/static/index.html");
});
// Route for serving the alumni.html form
app.get("/alumni", (req, res) => {
  res.sendFile(__dirname + "/static/alumni.html");
});

// Route to handle alumni form submission
app.post("/submit-alumni", async (req, res) => {
  try {
      const { Name, Email, phone, transactionID } = req.body;

      // Create CSV writer
      const csvWriter = createObjectCsvWriter({
          path: `${__dirname}/static/alumni_registrations.csv`,
          header: [
              { id: 'name', title: 'Name' },
              { id: 'email', title: 'Email' },
              { id: 'phone', title: 'Phone' },
              { id: 'transactionID', title: 'Transaction ID' }
          ],
          append: true // Append to the file if it exists
      });

      // Data to be written
      const data = [
          {
              name: Name,
              email: Email,
              phone: phone,
              transactionID: transactionID
          }
      ];

      // Write data to CSV file
      await csvWriter.writeRecords(data);

      // res.send("Registration successful! Your details have been saved.");
      res.sendFile(__dirname+ "/static/thenga2.html");
  } catch (error) {
      console.error('Error processing the form submission:', error);
      res.status(500).send('Internal Server Error');
  }
});

app.get("/ex-student", (req, res) => {
  res.sendFile(__dirname + "/static/form.html");
});

app.post("/submit", async (req, res) => {
  try {
    console.log(req.body);

    const name = req.body["Name"];
    const email = req.body["Email"];
    const no = req.body["phone"];
    const passOutYear = req.body["passoutyear"];
    const transactionID = req.body["transactionid"];

    // Generate a unique ID
    const uniqueId = uuidv4();

    // Save details to MongoDB
    const newRegistration = new Registration({
      name,
      email,
      phone: no,
      passOutYear,
      uniqueId,
      transactionID
    });

    await newRegistration.save();

    // Generate a QR code as a data URL
    const qrCodeData = `ID: ${uniqueId}, Name: ${name}, Email: ${email}, Phone: ${no}, Year of Pass Out: ${passOutYear}, Transation Id: ${transactionID}`;
    const qrCodeDataURL = await qrcode.toDataURL(qrCodeData);

    // Create a composite image with the QR code and user details
    const templateImagePath = __dirname + '/static/template.png';
    const templateImage = await loadImage(templateImagePath);
    const canvas = createCanvas(templateImage.width, templateImage.height);
    const ctx = canvas.getContext('2d');

    console.log("Template width " + templateImage.width + " and template height is " + templateImage.height);

    // Draw template
    ctx.drawImage(templateImage, 0, 0);

    // Draw QR code (enlarged)
    const qrCode = await loadImage(qrCodeDataURL);
    const qrCodeSize = 1000; 
    const qrX = (canvas.width - qrCodeSize) / 2;
    const qrY = (canvas.height / 2 - qrCodeSize) / 2 - 100; 
    ctx.drawImage(qrCode, qrX, qrY, qrCodeSize, qrCodeSize);

    // Add user details
    const textX = canvas.width / 2;
    const textY = canvas.height / 2 + 50; 
    ctx.textAlign = 'center';
    ctx.font = '70px Arial';
    ctx.fillStyle = '#000';
    ctx.fillText(`Name: ${name}`, textX, textY);
    ctx.fillText(`Email: ${email}`, textX, textY + 60);
    ctx.fillText(`Phone: ${no}`, textX, textY + 120);
    ctx.fillText(`Year of Pass Out: ${passOutYear}`, textX, textY + 180);
    ctx.fillText(`ID: ${uniqueId}`, textX, textY + 240);
    ctx.fillText(`Transaction ID: ${transactionID}`, textX, textY + 280);

    // Save composite image to file
    const outputImagePath = `${__dirname}/static/composite_${email}.png`;
    const buffer = canvas.toBuffer('image/png');
    await writeFile(outputImagePath, buffer);

    const htmlContent = `
      <html>
        <body>
          <p>Hello ${name},</p>
          <p>Thank you for submitting your information.</p>
          <p>We are so glad that you are coming!</p>
          <p>Attached is your unique QR code image.</p>
          <p>Best regards,<br>Your Company</p>
        </body>
      </html>`;

    // Configure the email transport using the default SMTP transport and a GMail account
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: 'bhattacharjee.agniva.jobs@gmail.com', 
        pass: 'dovz mfxv bfcy inpl'   
      }
    });

    // Email options with attachment
    const mailOptions = {
      from: 'bhattacharjee.agniva.jobs@gmail.com', 
      to: email,
      subject: 'Milan Pass!',
      html: htmlContent,
      attachments: [
        {
          filename: `composite_${email}.png`,
          path: outputImagePath,
          cid: 'unique@qr.code' // Same cid value as in the HTML img src
        }
      ]
    };

    // Send the email
    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        return console.log(error);
      }
      console.log('Email sent: ' + info.response);

      // Clean up the composite image file
      fs.unlink(outputImagePath, (err) => {
        if (err) {
          console.error('Error deleting the composite image file:', err);
        }
      });
    });

    res.sendFile(__dirname + "/static/thenga.html");
  } catch (error) {
    console.error('Error processing the form submission:', error);
    res.status(500).send('Internal Server Error');
  }
});

app.get('/download-excel', async (req, res) => {
  try {
    const registrations = await Registration.find().lean();

    const data = registrations.map(registration => [
      registration.name,
      registration.email,
      registration.phone,
      registration.passOutYear,
      registration.uniqueId,
      registration.transactionID,
      registration.entered ? 'Yes' : 'No'
    ]);

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Name", "Email", "Phone", "Pass Out Year", "Unique ID", "Transaction ID", "Entered"],
      ...data
    ]);

    XLSX.utils.book_append_sheet(workbook, worksheet, "Registrations");

    const excelFilePath = `${__dirname}/static/registrations.xlsx`;
    XLSX.writeFile(workbook, excelFilePath);

    res.download(excelFilePath, 'registrations.xlsx', err => {
      if (err) {
        console.error('Error downloading the file', err);
      }
      // Optional: Remove the file after download
      fs.unlinkSync(excelFilePath);
    });
  } catch (error) {
    console.error('Error generating the Excel file:', error);
    res.status(500).send('Internal Server Error');
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
