const express = require("express");
const dotenv = require("dotenv");
const fs = require('fs')
const xlsx = require('xlsx');
var bodyParser = require('body-parser')
// const data = require("./data.json")
var app = express()
const mysql = require("mysql");

const cron = require('node-cron')

const open = require('open');

// parse application/x-www-form-urlencoded
app.use(bodyParser.urlencoded({ extended : true }))

// parse application/json
app.use(bodyParser.json())
dotenv.config();

const nodemailer = require("nodemailer");

const cors = require("cors");
const corsOptions = {
  origin: "*",
  credentials: true, //access-control-allow-credentials:true
  optionSuccessStatus: 200,
};

app.use(cors()); // Use this after the variable declaration

app.use(express.json()); // tell the server to accept the json data from frontend



const multer = require('multer')

const storage = multer.diskStorage({
  destination: function(req, file, cb){
    return cb(null, "./public/Images")
  },
  filename: function (req, file, cb) {
    return cb(null, `${Date.now()}_${file.originalname}`)
  }
})

const upload = multer({storage})

let transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: process.env.SMTP_PORT,
  secure: false, // true for 465, false for other ports
  auth: {
    user: process.env.SMTP_MAIL, // generated ethereal user
    pass: process.env.SMTP_PASSWORD, // generated ethereal password
  },
});

const db = mysql.createConnection({
  host:'sql12.freesqldatabase.com',
  user:'sql12740281',
  password:'LRNYi97bju',
  database:'sql12740281'
});


app.post("/tryemail/:body/:type", upload.single('file') ,  async (req,res) =>{

  const details = JSON.parse(req.params.body)
  const type = req.params.type

  var mailOptions = {
    from: process.env.SMTP_MAIL,
    to: details.sendto,
    subject: details.subject,
    text: 
    type === 'FSI'?
    `From: https://fosconship.com/ \n 
    Name: ${details.name} \n 
    Phone number: ${details.phone} \n 
    Age: ${details.age} \n 
    Email: ${details.email} \n 
    Position/Rank: ${details.position} \n
    Year/s in Rank: ${details.years} \n 
    Type of Vessel: ${details.vessel} \n 
    Message: ${details.message} \n `
    :
    `From: https://fosconship.com/ \n 
    Name: ${details.name} \n 
    Phone number: ${details.phone} \n 
     Email: ${details.email} \n 
    Message: ${details.message} \n`
    ,
    attachments:[
      {
        filename: req.file.originalname,
        path:req.file.path
      }
    ]
  };
  console.log('sending')

  transporter.sendMail(mailOptions, function (error, info) {
    if (error) {
      console.log(error);
    } else {
      console.log("Email sent successfully!");
    setTimeout(()=>{
      fs.unlinkSync(req.file.path);
    },3000)
    res.send(200)
      }
  });
  
})


app.post("/hirings", upload.single('file') ,  (req, res) => {
  console.log(req.file)

  const workbook = xlsx.readFile(req.file.path) 
  // console.log(workbook)

  for(let indexZ = 0 ; indexZ <= workbook.SheetNames.length-1; indexZ++){
     
    const worksheet = workbook.Sheets[workbook.SheetNames[indexZ]]

  
    const range = xlsx.utils.decode_range(worksheet["!ref"])

    
                const deletes = "DELETE FROM hiringPosition"

                db.query(deletes,(err,result)=>{
                  if(err) return res.sendStatus(500)
        
                for(let row = range.s.r + 1 ; row <= range.e.r; row++){  

                    let data = []  
                        
                    try{
                        for(let col = range.s.c; col<=range.e.c; col++){
                            let cell = worksheet[xlsx.utils.encode_cell({r:row,c:col})]
                            data.push(cell.v)
                        }
                    }catch(err){
                    
                      return res.sendStatus(500)
                    }          

                      
                        const sqlpost = `INSERT INTO hiringPosition ( id , position, vesselType, remarks) VALUES (?,?,?,?);`

                        db.query(sqlpost,[row,data[0],data[1],data[2]], (err,result)=>{
                          if(err) return res.sendStatus(500)
                        })

                  
                }
                res.sendStatus(200)  
                })
    }     
   
}

);


app.get("/hiringPOS",  (req, res) => {

  const sqlget = "Select * FROM hiringPosition"

  db.query(sqlget,(err,result)=>{
    if(err) return res.sendStatus(500)
      // console.log(result)
    res.send(result);
   
})  

})


cron.schedule("10 * * * * *", async () => {
  console.log('awdawd')

  open('http://sindresorhus.com');

}, {
  timezone: "Asia/Manila"
})

app.get("/", function(req,res){

  res.send('helloworld')
})



const PORT = process.env.PORT;
app.listen(PORT, () => {
  console.log(`Example app listening on port ${PORT}`);
});
