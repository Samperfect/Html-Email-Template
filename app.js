const express = require('express');
const csv = require('fast-csv');
const { Nodemailing } = require('nodemailing');
const multer = require('multer');
const fs = require('fs');
const stream = require('stream');
const path = require('path');
const excel = require('xlsx');
require('dotenv').config();

// creating the app object
const app = express();

app.use(express.static('public'));
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));

const PORT = process.env.PORT || 8080;
var html = path.resolve(__dirname, 'public', 'email.html');
var row = path.resolve(__dirname, 'public', 'rows.html');

// importing the required modules
const googleCloud = require('@google-cloud/storage');
const multerGoogleStorage = require('multer-google-storage');

class Storage {
  constructor() {
    // setting up google firebase storage
    this.storage = new googleCloud.Storage({
      projectId: process.env.Firebase_Project_ID, //'<Firebase Project ID'
      keyFilename: process.env.Private_Key_JSON, //'<path to service accounts prviate key JSON>'
    });

    // setting up the firebase storage bucket
    this.bucket = this.storage.bucket(process.env.Order_Image_Bucket);

    // validating the file type of the images uploaded
    this.fileFilter = (req, file, cb) => {
      if (
        file.mimetype === 'image/jpeg' ||
        file.mimetype === 'image/png' ||
        file.mimetype === 'image/jpg'
      ) {
        cb(null, true);
      } else {
        cb(null, false);
      }
    };

    //image should not exceed 10 MB
    this.fileSize = 5 * 1024 * 1024;

    // setting up multer for form data handling
    this.upload = multer({
      //storing image as buffer in memory for use in firebase
      storage: multerGoogleStorage.storageEngine({
        keyFilename: process.env.Private_Key_JSON,
        projectId: process.env.Firebase_Project_ID,
        bucket: process.env.Order_Image_Bucket,
        acl: 'publicread',
        filename: (req, file, cb) => {
          const filename = this.generateFilename(file);
          cb(null, filename);
        },
      }),
      limits: { fileSize: this.fileSize }, //image should not exceed 10 MB
      // fileFilter: this.fileFilter,
    });
  }

  // helper function for creating new file name
  generateFilename(file) {
    return file.fieldname + '-' + Date.now() + '-' + file.originalname;
  }

  // helper function for deleting images from firestore
  deleteImages(files) {
    if (files) {
      files.map(async (file) => {
        try {
          await this.bucket.file(file.filename).delete();

          return true;
        } catch (error) {
          return false;
        }
      });
    }
  }
  async downloadImage(file) {
    if (file) {
      await this.bucket
        .file(file.filename)
        .download({ destination: __dirname + '/uploads/' + file.filename });

      return true;
    }
  }
}

let store = new Storage();

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

// function fabMails(sheet) {
//   return new Promise((resolve, reject) => {
//     try {
//       var output_file_name =
//         'https://emmyh-coin.appspot.com.storage.googleapis.com/mail.csv'; //path.resolve(__dirname, 'uploads', 'mail.csv');
//       var stream = excel.stream.to_csv(sheet);
//       stream.pipe(fs.createWriteStream(output_file_name));
//       resolve(output_file_name);
//     } catch (error) {
//       reject(error);
//     }
//   });
// }

app.get('/', (req, res) => {
  res.render('login');
  return;
});

app.post('/email', (req, res) => {
  if (!req.body) {
    res.redirect('/');
    return;
  }

  if (req.body.username === 'eton' && req.body.password === 'etong') {
    res.render('home');
    return;
  }

  res.redirect('/');
  return;
});

app.post('/upload', store.upload.any(), async (req, res) => {
  try {
    if (req.files == undefined) {
      return res.status(400).send({
        message: 'Please upload a CSV file!',
      });
    }
    await store.downloadImage(req.files[0]);

    // Import CSV File to MongoDB database
    let filePath = __dirname + '/uploads/' + req.files[0].filename;
    // let filePath = req.files[0].path;
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
      bannerImage2: ban[0].bannerImage2,
      bannerLink2: ban[0].bannerLink2,
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
  }
});

app.listen(PORT, () => {
  console.log('Now listening at port ' + PORT + '....');
});
