const express = require('express');
const csv = require('fast-csv');
const { Nodemailing } = require('nodemailing');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const excel = require('xlsx');
require('dotenv').config();

// creating the app object
const app = express();

app.use(express.static('public'));
app.set('view engine', 'ejs');

const PORT = process.env.PORT || 8080;
var html = path.resolve(__dirname, 'public', 'email.html');
var row = path.resolve(__dirname, 'public', 'rows.html');

// Multer Upload Storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname + '/uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, file.fieldname + '-' + Date.now() + '-' + file.originalname);
  },
});

// Filter for CSV file
const csvFilter = (req, file, cb) => {
  if (file.mimetype.includes('xlsx')) {
    cb(null, true);
  } else {
    cb('Please upload only csv file.', false);
  }
};
const upload = multer({ storage: storage });

app.get('/', (req, res) => {
  res.render('home');
});

app.post('/upload', upload.any(), async (req, res) => {
  try {
    if (req.files == undefined) {
      return res.status(400).send({
        message: 'Please upload a CSV file!',
      });
    }

    // Import CSV File to MongoDB database
    let filePath = __dirname + '/uploads/' + req.files[0].filename;
    let csvData = [];
    let file = excel.readFile(filePath);

    const sheets = file.SheetNames;

    const temp = excel.utils.sheet_to_json(file.Sheets[sheets[0]]);
    temp.forEach((res) => {
      csvData.push(res);
    });
    const ban = excel.utils.sheet_to_json(file.Sheets[sheets[1]]);

    const customers = excel.utils
      .sheet_to_json(file.Sheets[sheets[2]])
      .map(({ emails }) => emails);

    let product = ``;
    for (let i = 0; i < csvData.length; i++) {
      data = {
        name: csvData[i].name,
        price: csvData[i].price,
        discount: csvData[i].discount,
        image: csvData[i].image,
        link: csvData[i].link,
        name2: csvData[i].name2,
        price2: csvData[i].price2,
        discount2: csvData[i].discount2,
        image2: csvData[i].image2,
        link2: csvData[i].link2,
      };

      const prod = await Nodemailing.render(row, data);

      product += prod;
    }

    const final = await Nodemailing.render(html, {
      rows: product,
      bannerImage: ban[0].bannerImage,
      bannerLink: ban[0].bannerLink,
    });
    // 'queensong01@gmail.com',
    const msg = {
      Host: 'smtp.googlemail.com',
      Username: 'veehuelabs@gmail.com',
      Password: 'xrmfdbewnkerudsz',
      To: customers,
      From: 'etonng@gmail.com',
      Subject: `New Products to Buy`,
      Body: final,
    };

    Nodemailing.send(msg);

    res.redirect(
      `/?message=Your email has been successfully sent&class=alert-success`
    );
    return;

    // res.json({
    //   data: status,
    // });

    // fs.createReadStream(filePath)
    //   .pipe(csv.parse({ headers: true }))
    //   .on('error', (error) => {
    //     throw error.message;
    //   })
    //   .on('data', (row) => {
    //     csvData.push(row);
    //   })
    //   .on('end', async () => {
    //     let product = ``;
    //     for (let i = 0; i < csvData.length; i++) {
    //       data = {
    //         description: csvData[i].description,
    //         price: csvData[i].price,
    //         discount: csvData[i].discount,
    //         image: csvData[i].image,
    //         link: csvData[i].link,
    //       };

    //       const prod = await Nodemailing.render(row, data);

    //       product += prod;
    //     }

    //     const final = await Nodemailing.render(html, { row: product });

    //     // fs.writeFileSync(path.resolve(__dirname, 'public', 'test.html'), final);

    //     const bod = fs.readFileSync(
    //       path.resolve(__dirname, 'public', 'rows.html'),
    //       { encoding: 'utf8' }
    //     );

    //     const body = await Nodemailing.render(
    //       path.resolve(__dirname, 'public', 'email.html'),
    //       { row: bod }
    //     );

    //     const msg = {
    //       Host: 'smtp.googlemail.com',
    //       Username: 'veehuelabs@gmail.com',
    //       Password: 'xrmfdbewnkerudsz',
    //       To: 'lovedayperfection1@gmail.com',
    //       From: 'etonng@gmail.com',
    //       Subject: `New Products to Buy`,
    //       Body: body,
    //     };

    //     let status = await Nodemailing.send(msg);

    //     res.json({
    //       status: status,
    //     });
    //     return;
    //   });
  } catch (error) {
    // console.log('catch error-', error);
    res.redirect(
      `/?message=There was an error in the data you sent.&class=alert-danger`
    );
    return;
    // res.status(500).send({
    //   message: 'Could not upload the file: ' + req.file.originalname,
    // });
    return;
  }
});

app.listen(PORT, () => {
  console.log('Now listening at port ' + PORT + '....');
});
